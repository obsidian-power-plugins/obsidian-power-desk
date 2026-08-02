import { requestUrl } from "obsidian";
import { b64utf8 } from "./core";
import { PROPFIND_CALENDARS, PROPFIND_HOME, PROPFIND_PRINCIPAL, buildCalendarQuery, parseCalendarData, parseCollections, parseHomeSet, parsePrincipal, DavCollection } from "./caldavxml";

/* CalDAV transport: discovery (server URL to calendar collections) and the
 * time-ranged VEVENT REPORT. Obsidian's requestUrl passes PROPFIND/REPORT and
 * custom Depth headers through untouched and skips CORS on both desktop and
 * mobile, so this works everywhere with plain Basic auth (iCloud and Fastmail
 * hand out app-specific passwords for exactly this). */

export class CaldavError extends Error {
	constructor(
		message: string,
		readonly status?: number
	) {
		super(message);
	}
}

function friendly(status: number, server: string): string {
	if (status === 401 || status === 403) return `The CalDAV server refused the sign-in (${status}). Check the username and app password.`;
	if (status === 404) return `Nothing answers at ${server} (404). Check the server URL.`;
	return `The CalDAV server answered ${status}.`;
}

async function dav(url: string, method: string, depth: string, body: string, username: string, password: string): Promise<string> {
	const r = await requestUrl({
		url,
		method,
		headers: {
			Authorization: "Basic " + b64utf8(`${username}:${password}`),
			Depth: depth,
			"Content-Type": "application/xml; charset=utf-8",
		},
		body,
		throw: false,
	});
	if (r.status >= 400) throw new CaldavError(friendly(r.status, url), r.status);
	return r.text;
}

/** Server URL to its VEVENT calendar collections, via the standard three-step
 *  walk: current-user-principal, calendar-home-set, then the collections.
 *  Servers that skip a step (a home URL pasted directly) fall through, since
 *  each missing answer just reuses the previous URL. */
export async function discoverCalendars(serverUrl: string, username: string, password: string): Promise<DavCollection[]> {
	const base = serverUrl.trim().replace(/\/*$/, "/");
	let principal = base;
	try {
		const r = await dav(base, "PROPFIND", "0", PROPFIND_PRINCIPAL, username, password);
		principal = parsePrincipal(r, base) ?? base;
	} catch (e) {
		// iCloud's root answers PROPFIND; a server that 404s the root may still
		// serve calendars at the given URL directly, so only auth errors are fatal
		if (e instanceof CaldavError && (e.status === 401 || e.status === 403)) throw e;
	}
	let home = principal;
	try {
		const r = await dav(principal, "PROPFIND", "0", PROPFIND_HOME, username, password);
		home = parseHomeSet(r, principal) ?? principal;
	} catch (e) {
		if (e instanceof CaldavError && (e.status === 401 || e.status === 403)) throw e;
	}
	const r = await dav(home, "PROPFIND", "1", PROPFIND_CALENDARS, username, password);
	const cols = parseCollections(r, home);
	if (!cols.length) throw new CaldavError("Signed in, but no event calendars were found at this URL.");
	return cols;
}

/** Raw iCalendar payloads for events intersecting [startMs, endMs) in one
 *  collection. Each payload is a whole VCALENDAR (a series arrives with its
 *  overrides); the caller expands them via ics.ts. */
export async function fetchCollectionIcs(collectionHref: string, username: string, password: string, startMs: number, endMs: number): Promise<string[]> {
	const r = await dav(collectionHref, "REPORT", "1", buildCalendarQuery(startMs, endMs), username, password);
	return parseCalendarData(r);
}
