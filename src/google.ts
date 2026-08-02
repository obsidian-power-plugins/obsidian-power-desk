import { Platform, requestUrl } from "obsidian";

/* Google Calendar access. Auth is the loopback flow with PKCE against the
 * user's OWN Google Cloud project (client id + secret in settings): Google's
 * terms forbid shipping credentials inside an open-source plugin, its device
 * flow's scope allowlist excludes Calendar, and obsidian:// cannot be a
 * Google redirect URI, so a local port on desktop is the one clean path.
 * Sign-in is desktop-only; refresh and every API call work everywhere, and a
 * synced vault carries the connection to phones.
 *
 * Everything else is stateless REST over requestUrl, Graph-style: the plugin
 * owns token storage and refresh. */

export const GOOGLE_SCOPE = "openid email https://www.googleapis.com/auth/calendar";

const AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const TOKEN_URL = "https://oauth2.googleapis.com/token";
const API = "https://www.googleapis.com/calendar/v3";

export interface GoogleTokens {
	access_token: string;
	refresh_token?: string;
	expires_in: number;
	id_token?: string;
}

/** An error carrying the OAuth/API error code, so callers can tell a dead
 *  refresh token (reconnect) from a transient blip (retry). */
export class GoogleError extends Error {
	constructor(
		message: string,
		readonly code?: string
	) {
		super(message);
	}
}

function form(o: Record<string, string>): string {
	return Object.entries(o)
		.map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
		.join("&");
}

/** Obsidian's `RequestUrlResponse.json` throws on a non-JSON body; never let
 *  that mask the real status. */
function bodyJson(r: { json: unknown }): Record<string, unknown> | null {
	try {
		return (r.json as Record<string, unknown>) ?? null;
	} catch {
		return null;
	}
}

function apiError(r: { status: number; json: unknown }, doing: string): GoogleError {
	const err = bodyJson(r)?.error as { code?: number; message?: string; errors?: { reason?: string }[]; status?: string } | string | undefined;
	if (typeof err === "string") return new GoogleError(`Could not ${doing} (${err}).`, err);
	return new GoogleError(err?.message || `Could not ${doing} (${r.status}).`, err?.errors?.[0]?.reason || err?.status);
}

const b64url = (bytes: Uint8Array) => {
	let bin = "";
	for (const b of bytes) bin += String.fromCharCode(b);
	return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
};

/** Loopback sign-in: start a one-shot server on 127.0.0.1, send the user's
 *  browser to Google, catch the code, exchange it. Resolves with tokens or
 *  rejects with a readable error; always tears the server down. */
export function loopbackAuth(clientId: string, clientSecret: string, openBrowser: (url: string) => void): Promise<GoogleTokens> {
	if (!Platform.isDesktopApp) {
		return Promise.reject(new GoogleError("Google sign-in needs the desktop app once; the connection then syncs to every device."));
	}
	const http = require("node:http") as typeof import("node:http");
	const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
	return new Promise<GoogleTokens>((resolve, reject) => {
		let redirect = "";
		let settled = false;
		let timer = 0;
		const settle = (fn: () => void) => {
			if (settled) return;
			settled = true;
			window.clearTimeout(timer);
			try {
				server.close();
			} catch {
				/* already down */
			}
			fn();
		};
		const server = http.createServer((req, res) => {
			const u = new URL(req.url ?? "/", "http://127.0.0.1");
			const code = u.searchParams.get("code");
			const err = u.searchParams.get("error");
			res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
			res.end("<html><body style=\"font-family: sans-serif; padding: 2em\">Signed in. You can close this window and return to Obsidian.</body></html>");
			if (err) settle(() => reject(new GoogleError(`Google refused the sign-in (${err}).`, err)));
			else if (code) settle(() => void exchangeCode(clientId, clientSecret, code, redirect, verifier).then(resolve, reject));
		});
		server.on("error", (e) => settle(() => reject(new GoogleError(`Could not open a local sign-in port (${String(e)}).`))));
		server.listen(0, "127.0.0.1", () => {
			const addr = server.address();
			const port = addr && typeof addr === "object" ? addr.port : 0;
			redirect = `http://127.0.0.1:${port}`;
			void crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)).then((digest) => {
				const url = `${AUTH_URL}?${form({
					client_id: clientId,
					redirect_uri: redirect,
					response_type: "code",
					scope: GOOGLE_SCOPE,
					access_type: "offline",
					prompt: "consent",
					code_challenge: b64url(new Uint8Array(digest)),
					code_challenge_method: "S256",
				})}`;
				openBrowser(url);
			});
		});
		timer = window.setTimeout(() => settle(() => reject(new GoogleError("Sign-in timed out; try again."))), 5 * 60_000);
	});
}

async function exchangeCode(clientId: string, clientSecret: string, code: string, redirect: string, verifier: string): Promise<GoogleTokens> {
	const r = await requestUrl({
		url: TOKEN_URL,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: form({ grant_type: "authorization_code", code, client_id: clientId, client_secret: clientSecret, redirect_uri: redirect, code_verifier: verifier }),
		throw: false,
	});
	if (r.status >= 400) {
		const j = bodyJson(r);
		throw new GoogleError((j?.error_description as string) || `Sign-in failed (${(j?.error as string) || r.status}).`, j?.error as string | undefined);
	}
	return r.json as GoogleTokens;
}

export async function refreshGoogleTokens(clientId: string, clientSecret: string, refreshToken: string): Promise<GoogleTokens> {
	const r = await requestUrl({
		url: TOKEN_URL,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: form({ grant_type: "refresh_token", client_id: clientId, client_secret: clientSecret, refresh_token: refreshToken }),
		throw: false,
	});
	if (r.status >= 400) {
		const j = bodyJson(r);
		// a coded refusal (invalid_grant, ...) means the refresh token is dead;
		// no code reads as a transient network failure worth keeping tokens for
		throw new GoogleError((j?.error_description as string) || `Google session expired; reconnect (${r.status}).`, j?.error as string | undefined);
	}
	return r.json as GoogleTokens;
}

async function pagedGet(baseUrl: string, accessToken: string, itemsKey: string, doing: string, maxPages = 20): Promise<unknown[]> {
	const out: unknown[] = [];
	let pageToken: string | null = null;
	for (let page = 0; page < maxPages; page++) {
		const url = pageToken ? `${baseUrl}&pageToken=${encodeURIComponent(pageToken)}` : baseUrl;
		const r = await requestUrl({ url, method: "GET", headers: { Authorization: `Bearer ${accessToken}` }, throw: false });
		if (r.status >= 400) throw apiError(r, doing);
		const j = bodyJson(r);
		out.push(...((j?.[itemsKey] as unknown[]) ?? []));
		pageToken = (j?.nextPageToken as string | undefined) ?? null;
		if (!pageToken) break;
	}
	return out;
}

export interface GoogleCalendarInfo {
	id: string;
	name: string;
	color: string;
	primary: boolean;
	writable: boolean;
}

export async function listGoogleCalendars(accessToken: string): Promise<GoogleCalendarInfo[]> {
	const raw = await pagedGet(`${API}/users/me/calendarList?maxResults=250&showHidden=false`, accessToken, "items", "list your calendars");
	return raw
		.map((c) => {
			const o = c as Record<string, unknown>;
			return {
				id: String(o.id ?? ""),
				name: String(o.summaryOverride ?? o.summary ?? "Calendar"),
				color: typeof o.backgroundColor === "string" && /^#[0-9a-fA-F]{6}$/.test(o.backgroundColor) ? o.backgroundColor.toLowerCase() : "",
				primary: !!o.primary,
				writable: o.accessRole === "owner" || o.accessRole === "writer",
			};
		})
		.filter((c) => c.id);
}

/** Concrete instances between two ISO instants (singleEvents expands series
 *  server-side, exactly like Graph's calendarView). */
export async function fetchGoogleEvents(accessToken: string, calendarId: string, startISO: string, endISO: string): Promise<unknown[]> {
	const url =
		`${API}/calendars/${encodeURIComponent(calendarId)}/events?singleEvents=true&orderBy=startTime&maxResults=250` +
		`&timeMin=${encodeURIComponent(startISO)}&timeMax=${encodeURIComponent(endISO)}`;
	return pagedGet(url, accessToken, "items", "read the calendar");
}

/** Google only notifies guests when asked: sendUpdates defaults to none at
 *  the API layer, so writes involving people must opt in. */
const notifyQs = (notify: boolean) => (notify ? "?sendUpdates=all" : "");

export async function insertGoogleEvent(accessToken: string, calendarId: string, body: Record<string, unknown>, notify = false): Promise<void> {
	const r = await requestUrl({
		url: `${API}/calendars/${encodeURIComponent(calendarId)}/events${notifyQs(notify)}`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw apiError(r, "create the event");
}

/** Patch an event or one expanded instance (Google turns an instance patch
 *  into an exception, like Graph). */
export async function patchGoogleEvent(accessToken: string, calendarId: string, eventId: string, body: Record<string, unknown>, notify = false): Promise<void> {
	const r = await requestUrl({
		url: `${API}/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(eventId)}${notifyQs(notify)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw apiError(r, "update the event");
}

/** One event by id, raw; RSVP needs the full attendee array to patch back. */
export async function getGoogleEvent(accessToken: string, calendarId: string, eventId: string): Promise<Record<string, unknown>> {
	const r = await requestUrl({
		url: `${API}/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(eventId)}`,
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw apiError(r, "read the event");
	return (r.json as Record<string, unknown>) ?? {};
}

export async function deleteGoogleEvent(accessToken: string, calendarId: string, eventId: string): Promise<void> {
	const r = await requestUrl({
		url: `${API}/calendars/${encodeURIComponent(calendarId)}/events/${encodeURIComponent(eventId)}?sendUpdates=all`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400 && r.status !== 410) throw apiError(r, "delete the event"); // 410: already gone, which is what deleting wanted
}
