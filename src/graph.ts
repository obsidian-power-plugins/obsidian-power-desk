import { arrayBufferToBase64, requestUrl } from "obsidian";
import { buildMessageBatch, parseMessageBatch, buildCategoryPatchBatch, buildReadPatchBatch, parseWriteBatch } from "./core";
import type { BatchOutcome } from "./core";
import { wallOfMs, msOfKey, keyOfMs } from "./core";

/* Microsoft Graph access: OAuth device-code flow (the user signs in themselves
 * in a browser, so no password ever touches the plugin) and read-only calendar
 * fetches. Stateless HTTP only; the plugin owns token storage and refresh.
 * Requires an Azure app registration whose client id the user provides in
 * settings, configured as a public client.
 *
 * The device-code flow is the one Graph sign-in that works identically on
 * desktop and mobile: a POST plus polling, no loopback listener, no redirect
 * URI, no Node APIs. It is why this plugin ships without a desktop-only flag. */

/** The scope this version asks for. Connections made under an older, narrower
 *  scope keep refreshing against what they were granted (see the accounts'
 *  grantedScope), so an un-reconnected session keeps its old powers instead of
 *  dying with invalid_grant the first time it refreshes; the new ones arrive
 *  with one reconnect. */
// MailboxSettings.ReadWrite is what inbox rules live under: Mail.ReadWrite
// does not reach them, however much it sounds like it should.
export const GRAPH_SCOPE = "offline_access openid profile Calendars.ReadWrite Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite ProfilePhoto.Read.All Contacts.Read Tasks.ReadWrite";
export const GRAPH_READ_SCOPE = "offline_access openid profile Calendars.Read";
const SCOPE = GRAPH_SCOPE;

export interface DeviceCode {
	device_code: string;
	user_code: string;
	verification_uri: string;
	interval: number;
	expires_in: number;
	message: string;
}

export interface GraphTokens {
	access_token: string;
	refresh_token: string;
	expires_in: number;
	/** Present because SCOPE asks for openid; its claims label the account. */
	id_token?: string;
}

const authBase = (tenant: string) => `https://login.microsoftonline.com/${encodeURIComponent(tenant || "common")}/oauth2/v2.0`;

function form(o: Record<string, string>): string {
	return Object.entries(o)
		.map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
		.join("&");
}

/** Obsidian's `RequestUrlResponse.json` is a lazy getter that THROWS on a
 *  non-JSON body (an HTML 502, an empty 429). Never let that mask the real
 *  status; return null and let callers treat it as transient. */
function bodyJson(r: { json: unknown }): Record<string, unknown> | null {
	try {
		return (r.json as Record<string, unknown>) ?? null;
	} catch {
		return null;
	}
}

/** An error carrying the OAuth error code, so the caller can tell a rejected
 *  refresh token (reconnect) from a transient network blip (retry). */
export class GraphError extends Error {
	constructor(
		message: string,
		readonly code?: string
	) {
		super(message);
	}
}

/** Begin the device-code flow: returns a code and URL to show the user. */
export async function startDeviceCode(clientId: string, tenant: string): Promise<DeviceCode> {
	const r = await requestUrl({
		url: `${authBase(tenant)}/devicecode`,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: form({ client_id: clientId, scope: SCOPE }),
		throw: false,
	});
	if (r.status >= 400) throw new Error((bodyJson(r)?.error_description as string) || `Could not start sign-in (${r.status}).`);
	return r.json as DeviceCode;
}

/** Poll once for tokens. Returns "pending" while the user is still signing in. */
export async function pollToken(clientId: string, tenant: string, deviceCode: string): Promise<GraphTokens | "pending"> {
	const r = await requestUrl({
		url: `${authBase(tenant)}/token`,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: form({ grant_type: "urn:ietf:params:oauth:grant-type:device_code", client_id: clientId, device_code: deviceCode }),
		throw: false,
	});
	if (r.status < 400) return r.json as GraphTokens;
	const j = bodyJson(r);
	if (!j) return "pending"; // a non-JSON 429/5xx blip: keep polling rather than abort sign-in
	const err = j.error as string | undefined;
	if (err === "authorization_pending" || err === "slow_down") return "pending";
	throw new Error((j.error_description as string) || `Sign-in failed (${err || r.status}).`);
}

/** Exchange a refresh token for a fresh access token (and rotated refresh).
 *  `scope` must be what the token was consented for, not what we wish it were. */
export async function refreshTokens(clientId: string, tenant: string, refreshToken: string, scope: string = SCOPE): Promise<GraphTokens> {
	const r = await requestUrl({
		url: `${authBase(tenant)}/token`,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: form({ grant_type: "refresh_token", client_id: clientId, refresh_token: refreshToken, scope }),
		throw: false,
	});
	if (r.status >= 400) {
		const j = bodyJson(r);
		// a server-returned code (invalid_grant, ...) means the refresh token is
		// dead; no code means a transient network failure worth keeping tokens for
		throw new GraphError((j?.error_description as string) || `Session expired; reconnect Microsoft 365 (${r.status}).`, j?.error as string | undefined);
	}
	return r.json as GraphTokens;
}

/** GET a Graph collection and follow @odata.nextLink until it runs dry. A
 *  busy calendar over a six-week month grid overflows one page easily; a
 *  hard page cap keeps a pathological mailbox from spinning forever. */
async function pagedGet(url: string, accessToken: string, extraHeaders: Record<string, string> = {}, maxPages = 20): Promise<unknown[]> {
	const out: unknown[] = [];
	let next: string | null = url;
	for (let page = 0; next && page < maxPages; page++) {
		const r = await requestUrl({
			url: next,
			method: "GET",
			headers: { Authorization: `Bearer ${accessToken}`, ...extraHeaders },
			throw: false,
		});
		if (r.status >= 400) {
			const j = bodyJson(r);
			const err = j?.error as { code?: string; message?: string } | undefined;
			throw new GraphError(err?.message || `Could not read your calendar (${r.status}).`, err?.code);
		}
		const j = bodyJson(r);
		out.push(...((j?.value as unknown[]) ?? []));
		next = (j?.["@odata.nextLink"] as string | undefined) ?? null;
	}
	return out;
}

export interface GraphCalendarInfo {
	id: string;
	name: string;
	hexColor: string;
	isDefaultCalendar: boolean;
}

/** The signed-in user's calendars (name, color, default flag). */
export async function listCalendars(accessToken: string): Promise<GraphCalendarInfo[]> {
	const raw = await pagedGet("https://graph.microsoft.com/v1.0/me/calendars?$select=id,name,hexColor,isDefaultCalendar&$top=100", accessToken);
	return raw
		.map((c) => {
			const o = c as Record<string, unknown>;
			return {
				id: String(o.id ?? ""),
				name: String(o.name ?? "Calendar"),
				hexColor: typeof o.hexColor === "string" ? o.hexColor : "",
				isDefaultCalendar: !!o.isDefaultCalendar,
			};
		})
		.filter((c) => c.id);
}

const EVENT_SELECT = "id,subject,start,end,isAllDay,isCancelled,showAs,type,location,organizer,attendees,onlineMeeting,isOnlineMeeting,webLink,bodyPreview,responseStatus,isOrganizer,seriesMasterId,categories";

/** Concrete event instances between two ISO instants, in the device's own
 *  timezone (calendarView expands recurring series server-side). Pass a
 *  calendarId for a specific calendar, null for the default one. */
export async function fetchCalendarView(accessToken: string, calendarId: string | null, startISO: string, endISO: string): Promise<unknown[]> {
	const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
	const base = calendarId ? `https://graph.microsoft.com/v1.0/me/calendars/${encodeURIComponent(calendarId)}/calendarView` : "https://graph.microsoft.com/v1.0/me/calendarView";
	const url =
		`${base}?startDateTime=${encodeURIComponent(startISO)}&endDateTime=${encodeURIComponent(endISO)}` +
		`&$orderby=${encodeURIComponent("start/dateTime")}&$top=250&$select=${EVENT_SELECT}`;
	// the Prefer header rides along on every page, so paged results stay in one zone
	return pagedGet(url, accessToken, { Prefer: `outlook.timezone="${tz}"` });
}

/* ---------------- mail (triage inbox) ---------------- */

// conversationId threads a back-and-forth into one exchange; inferenceClassification
// is Outlook's own focused/other relevance verdict, already personalized to this
// mailbox. Both are plain properties, so they cost nothing extra on the list.
const MAIL_SELECT =
	"id,subject,from,receivedDateTime,bodyPreview,isRead,webLink,hasAttachments,parentFolderId,importance,flag,toRecipients,conversationId,inferenceClassification,categories";

/** The newest messages of one folder ("inbox" by well-known name, or any
 *  folder id), one page; triage wants recency, not archives. */
export async function fetchFolderMessages(accessToken: string, folderId = "inbox", top = 25): Promise<unknown[]> {
	const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages?$top=${top}&$select=${MAIL_SELECT}&$orderby=${encodeURIComponent("receivedDateTime desc")}`;
	return pagedGet(url, accessToken, {}, 1);
}

/** Every message in a folder back to `sinceMs`, newest first, for a one-time
 *  import rather than triage. Pages as deep as `cap` asks; the caller bounds
 *  it, because a work folder can hold years. */
export async function fetchFolderMessagesDeep(accessToken: string, folderId: string, sinceMs: number, cap = 2000): Promise<unknown[]> {
	const filter = encodeURIComponent(`receivedDateTime ge ${new Date(sinceMs).toISOString()}`);
	const url =
		`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages` +
		`?$top=50&$select=${MAIL_SELECT}&$filter=${filter}&$orderby=${encodeURIComponent("receivedDateTime desc")}`;
	const pages = Math.min(80, Math.max(1, Math.ceil(cap / 50)));
	const all = await pagedGet(url, accessToken, {}, pages);
	return all.slice(0, cap);
}

export interface DeltaResult {
	messages: unknown[];
	removedIds: string[];
	/** Null when the round could not finish (page cap) and the next refresh
	 *  should start over. */
	deltaLink: string | null;
	/** The server discarded our token (410): drop the cached list and resync. */
	resync?: boolean;
}

/** One delta round for a folder's messages: with no link, an initial sync
 *  bounded to the recent window; with one, only what changed since last time,
 *  including removals. This is the mechanism that keeps folder lists current
 *  without re-downloading them, the way Outlook's cached mode stays fast. */
export async function deltaFolderMessages(accessToken: string, folderId: string, deltaLink: string | null, sinceMs: number, maxMessages = 500): Promise<DeltaResult> {
	const filter = encodeURIComponent(`receivedDateTime ge ${new Date(sinceMs).toISOString()}`);
	let url = deltaLink ?? `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages/delta?$select=${MAIL_SELECT}&$filter=${filter}`;
	const messages: unknown[] = [];
	const removedIds: string[] = [];
	// page deep enough to fill the caller's target (50 per page), with a hard
	// ceiling so a runaway link can never loop forever
	const maxPages = Math.min(80, Math.max(10, Math.ceil(maxMessages / 50) + 1));
	for (let page = 0; page < maxPages && url; page++) {
		const r = await requestUrl({
			url,
			method: "GET",
			headers: { Authorization: `Bearer ${accessToken}`, Prefer: "odata.maxpagesize=50" },
			throw: false,
		});
		if (r.status === 410) return { messages: [], removedIds: [], deltaLink: null, resync: true };
		if (r.status >= 400) throw writeError(r, "sync the folder");
		const j = bodyJson(r);
		for (const item of (j?.value as Record<string, unknown>[] | undefined) ?? []) {
			if (item["@removed"]) {
				if (typeof item.id === "string") removedIds.push(item.id);
			} else messages.push(item);
		}
		const delta = j?.["@odata.deltaLink"];
		if (typeof delta === "string") return { messages, removedIds, deltaLink: delta };
		const next = j?.["@odata.nextLink"];
		url = typeof next === "string" ? next : "";
	}
	return { messages, removedIds, deltaLink: null };
}

/** Mailbox-wide search, Outlook style: subject, sender, and body all match.
 *  Graph forbids $orderby beside $search (results come relevance-ranked), so
 *  the caller re-sorts by date. */
export async function searchMessages(accessToken: string, query: string, top = 25): Promise<unknown[]> {
	const q = encodeURIComponent(`"${query.replace(/"/g, "")}"`);
	const url = `https://graph.microsoft.com/v1.0/me/messages?$search=${q}&$top=${top}&$select=${MAIL_SELECT}`;
	return pagedGet(url, accessToken, {}, 1);
}

/** The newest unread messages across the WHOLE mailbox (Outlook's Unread
 *  Mail search folder); the caller scopes them to the inbox subtree with the
 *  folder ids each message carries. The dummy receivedDateTime bound exists
 *  because Graph requires every $orderby property to appear first in the
 *  $filter when both are used. */
export async function fetchUnreadMessages(accessToken: string, maxPages = 2): Promise<unknown[]> {
	const filter = encodeURIComponent("receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false");
	const url = `https://graph.microsoft.com/v1.0/me/messages?$filter=${filter}&$orderby=${encodeURIComponent("receivedDateTime desc")}&$top=50&$select=${MAIL_SELECT}`;
	return pagedGet(url, accessToken, {}, maxPages);
}

const FOLDER_SELECT = "id,displayName,parentFolderId,childFolderCount,unreadItemCount,totalItemCount";

/** The inbox's folder id, so tree ordering never depends on a display name
 *  that localization could rename. */
export async function getInboxId(accessToken: string): Promise<string | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=id",
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) return null;
	const id = bodyJson(r)?.id;
	return typeof id === "string" ? id : null;
}

/** The whole folder tree: the top level, then children breadth-first (only
 *  where childFolderCount says there are any), a few levels deep. */
export async function fetchMailFolders(accessToken: string): Promise<unknown[]> {
	const base = "https://graph.microsoft.com/v1.0/me/mailFolders";
	const out: unknown[] = [];
	const frontier: { id: string | null; depth: number }[] = [{ id: null, depth: 0 }];
	while (frontier.length) {
		const next = frontier.shift();
		if (!next) break;
		const url = next.id
			? `${base}/${encodeURIComponent(next.id)}/childFolders?$top=100&$select=${FOLDER_SELECT}`
			: `${base}?$top=100&$select=${FOLDER_SELECT}`;
		const page = await pagedGet(url, accessToken, {}, 2);
		out.push(...page);
		if (next.depth < 3) {
			for (const f of page) {
				const o = f as Record<string, unknown>;
				if (typeof o.id === "string" && typeof o.childFolderCount === "number" && o.childFolderCount > 0) {
					frontier.push({ id: o.id, depth: next.depth + 1 });
				}
			}
		}
	}
	return out;
}

/** One message with its full body, for the reading pane. */
export interface MailAttachment {
	id: string;
	name: string;
	contentType: string;
	size: number;
	isInline: boolean;
	contentId?: string;
}

/** The attachment list without payloads; bytes come per attachment on demand. */
export async function listMailAttachments(accessToken: string, messageId: string): Promise<MailAttachment[]> {
	const r = await requestUrl({
		// only base attachment properties: contentId lives on the fileAttachment
		// subtype and naming it here fails the whole request with a 400
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline&$top=50`,
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "list the attachments");
	return ((r.json as { value?: Record<string, unknown>[] }).value ?? []).map((v) => ({
		id: String(v.id ?? ""),
		name: String(v.name ?? "attachment"),
		contentType: String(v.contentType ?? "application/octet-stream"),
		size: Number(v.size ?? 0),
		isInline: !!v.isInline,
		contentId: typeof v.contentId === "string" ? v.contentId : undefined,
	}));
}

/** One attachment with its base64 payload; null for item or reference
 *  attachments, which carry no bytes. */
export async function getMailAttachmentBytes(accessToken: string, messageId: string, attachmentId: string): Promise<{ name: string; contentType: string; contentBytes: string; contentId?: string } | null> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "download the attachment");
	const v = (r.json as Record<string, unknown>) ?? {};
	if (typeof v.contentBytes !== "string") return null;
	return {
		name: String(v.name ?? "attachment"),
		contentType: String(v.contentType ?? "application/octet-stream"),
		contentBytes: v.contentBytes,
		contentId: typeof v.contentId === "string" ? v.contentId : undefined,
	};
}

export async function getMessage(accessToken: string, messageId: string): Promise<Record<string, unknown>> {
	const r = await requestUrl({
		// internetMessageHeaders rides along so the unsubscribe links cost no
		// request of their own: the body is already being fetched
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}?$select=subject,from,toRecipients,receivedDateTime,body,webLink,internetMessageHeaders`,
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "read the message");
	return (r.json as Record<string, unknown>) ?? {};
}

/** Read many messages in one round trip via Graph's $batch endpoint.
 *
 *  The reason this exists: fetching bodies one message at a time is an N+1,
 *  and a ninety-day backfill is thousands of sequential round trips. Twenty
 *  per call is Graph's ceiling, so this is a ~20x reduction in requests for
 *  the same data over a fully supported API.
 *
 *  A sub-request that fails does not fail its neighbours; the caller gets the
 *  ids back and decides whether to retry. */
export async function getMessagesBatch(
	accessToken: string,
	ids: readonly string[],
	select = "subject,from,toRecipients,receivedDateTime,body,webLink"
): Promise<BatchOutcome> {
	if (!ids.length) return { ok: new Map(), failed: [], retryAfterMs: 0 };
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/$batch",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(buildMessageBatch(ids, select)),
		throw: false,
	});
	// a whole-batch failure is every id failed, not an exception: the caller is
	// usually prefetching and a warm cache is best-effort by nature
	if (r.status >= 400) {
		const secs = parseFloat(String((r.headers as Record<string, string> | undefined)?.["retry-after"] ?? "0"));
		return { ok: new Map(), failed: [...ids], retryAfterMs: isFinite(secs) && secs > 0 ? secs * 1000 : 0 };
	}
	return parseMessageBatch(r.json, ids);
}

export async function markMessageRead(accessToken: string, messageId: string, read: boolean): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ isRead: read }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "update the message");
}

/** Flag or clear a message for follow-up.
 *
 *  Graph models the flag as a status rather than a boolean, and "notFlagged"
 *  is the one that clears it; PATCHing the flag away with null leaves a
 *  completed flag behind instead of no flag at all. */
export async function flagMessage(accessToken: string, messageId: string, flagged: boolean): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ flag: { flagStatus: flagged ? "flagged" : "notFlagged" } }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "flag the message");
}

/** Move to the well-known Archive folder, the one-key triage verb. */
export async function archiveMessage(accessToken: string, messageId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/move`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ destinationId: "archive" }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "archive the message");
}

/** Move a message into any folder by id, which is what snooze needs in both
 *  directions: out to the holding folder and back where it came from. */
export async function moveMessage(accessToken: string, messageId: string, destinationId: string): Promise<string | null> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/move`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ destinationId }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "move the message");
	// a move mints a new id in the destination folder; the old one is gone,
	// so anything that means to touch the message again has to carry this
	return (bodyJson(r)?.id as string) ?? null;
}

/** Make a folder, at the mailbox root or inside another. */
export async function createMailFolder(accessToken: string, displayName: string, parentId?: string | null): Promise<string | null> {
	const url = parentId
		? `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(parentId)}/childFolders`
		: "https://graph.microsoft.com/v1.0/me/mailFolders";
	const r = await requestUrl({
		url,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ displayName }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, `create the ${displayName} folder`);
	return (bodyJson(r)?.id as string) ?? null;
}

export async function renameMailFolder(accessToken: string, folderId: string, displayName: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ displayName }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "rename the folder");
}

export async function deleteMailFolder(accessToken: string, folderId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the folder");
}

/** Find a folder by display name at the mailbox root, creating it if it is
 *  not there. Snooze needs somewhere to park mail that is not the inbox and
 *  not Archive, and a real Outlook folder means the parked mail is visible
 *  and recoverable from any client, not hidden in this plugin's head. */
export async function ensureMailFolder(accessToken: string, displayName: string): Promise<string | null> {
	const find = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=id,displayName`,
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (find.status < 400) {
		const list = (bodyJson(find)?.value as { id?: string; displayName?: string }[] | undefined) ?? [];
		const hit = list.find((f) => (f.displayName ?? "").toLowerCase() === displayName.toLowerCase());
		if (hit?.id) return hit.id;
	}
	return createMailFolder(accessToken, displayName);
}

/** A draft built from scratch, for the compose window's scheduled sends.
 *  Immediate sends still go through sendMail, which needs no draft at all;
 *  a deferred one has to exist as a message before the send time can be
 *  hung off it. */
export async function createDraftMessage(accessToken: string, m: { to: string[]; cc?: string[]; bcc?: string[]; subject: string; html: string; receipts?: Receipts }): Promise<string | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/messages",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			subject: m.subject,
			body: { contentType: "HTML", content: m.html },
			toRecipients: m.to.map((a) => ({ emailAddress: { address: a } })),
			ccRecipients: (m.cc ?? []).map((a) => ({ emailAddress: { address: a } })),
			bccRecipients: (m.bcc ?? []).map((a) => ({ emailAddress: { address: a } })),
			...receiptFields(m.receipts),
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "create the draft");
	return (bodyJson(r)?.id as string) ?? null;
}

/** Graph takes an attachment in one request up to about 3 MB; past that it
 *  wants an upload session. The cutoff is set below the limit because the
 *  base64 body is a third larger than the file it carries. */
const SIMPLE_ATTACH_MAX = 3 * 1024 * 1024;
/** Upload chunks must be a multiple of 320 KiB. Nine of them is a shade
 *  under 3 MB, which keeps each request comfortably inside the limit. */
const UPLOAD_CHUNK = 327680 * 9;

export interface OutgoingFile {
	name: string;
	contentType: string;
	bytes: ArrayBuffer;
}

/** Attach a file to a draft, in one request or in chunks depending on size.
 *
 *  The chunked path uploads to a pre-authenticated URL Graph hands back, and
 *  that URL must NOT carry the Authorization header: it is already scoped to
 *  this one upload, and sending the bearer token alongside it is rejected. */
/** Attach an image the body refers to by content id, so it renders in place
 *  rather than as a file at the bottom. Always small enough for one request:
 *  a signature logo that needed a chunked upload would be a problem of its
 *  own. */
export async function addInlineImage(accessToken: string, draftId: string, img: { cid: string; name: string; contentType: string; base64: string }): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}/attachments`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			"@odata.type": "#microsoft.graph.fileAttachment",
			name: img.name,
			contentType: img.contentType,
			contentBytes: img.base64,
			isInline: true,
			contentId: img.cid,
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, `attach the image ${img.name}`);
}

export async function addFileAttachment(accessToken: string, draftId: string, file: OutgoingFile): Promise<void> {
	const size = file.bytes.byteLength;
	if (size <= SIMPLE_ATTACH_MAX) {
		const r = await requestUrl({
			url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}/attachments`,
			method: "POST",
			contentType: "application/json",
			headers: { Authorization: `Bearer ${accessToken}` },
			body: JSON.stringify({
				"@odata.type": "#microsoft.graph.fileAttachment",
				name: file.name,
				contentType: file.contentType || "application/octet-stream",
				contentBytes: arrayBufferToBase64(file.bytes),
			}),
			throw: false,
		});
		if (r.status >= 400) throw writeError(r, `attach ${file.name}`);
		return;
	}

	const open = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}/attachments/createUploadSession`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ AttachmentItem: { attachmentType: "file", name: file.name, size, contentType: file.contentType || "application/octet-stream" } }),
		throw: false,
	});
	if (open.status >= 400) throw writeError(open, `start the upload for ${file.name}`);
	const uploadUrl = bodyJson(open)?.uploadUrl as string | undefined;
	if (!uploadUrl) throw new Error(`Power Desk could not start the upload for ${file.name}.`);

	for (let start = 0; start < size; start += UPLOAD_CHUNK) {
		const end = Math.min(start + UPLOAD_CHUNK, size);
		const chunk = await requestUrl({
			url: uploadUrl,
			method: "PUT",
			headers: { "Content-Range": `bytes ${start}-${end - 1}/${size}`, "Content-Length": String(end - start) },
			body: file.bytes.slice(start, end),
			throw: false,
		});
		// 200 and 201 close the upload, 202 asks for the next range
		if (chunk.status >= 400) throw writeError(chunk, `upload ${file.name}`);
	}
}

/** Hang a send time on a draft: PidTagDeferredSendTime, the MAPI property
 *  Outlook's own Delay Delivery sets.
 *
 *  This is the reason schedule send works at all here. Transport holds the
 *  message and releases it at the time, so it goes out whether or not
 *  Obsidian, or the computer, is still on. Everything else about "later" in
 *  a plugin has to wait for the app to be open; this one does not. The value
 *  is ISO 8601, and Graph wants it in UTC. */
export async function setDeferredSend(accessToken: string, draftId: string, whenMs: number): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			singleValueExtendedProperties: [{ id: "SystemTime 0x3FEF", value: new Date(whenMs).toISOString() }],
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "schedule the message");
}

/** The one-click unsubscribe POST of RFC 8058.
 *
 *  Deliberately naked: no Authorization header, no cookies, no mailbox
 *  identity. The URL already carries whatever token the sender needs, and
 *  this is a request to a stranger's server, so it gets nothing of yours
 *  beyond the body the standard specifies. */
export async function postOneClickUnsubscribe(url: string): Promise<void> {
	const r = await requestUrl({
		url,
		method: "POST",
		contentType: "application/x-www-form-urlencoded",
		body: "List-Unsubscribe=One-Click",
		throw: false,
	});
	if (r.status >= 400) throw new Error(`The sender's server answered ${r.status}.`);
}

/** Report a message as junk: the sender joins the blocked list and the
 *  message moves to Junk Email, which is what teaches the filter.
 *
 *  This is the v1.0 action. Submitting a message to Microsoft's security
 *  team, which is what Outlook's Report phishing does on top of this, lives
 *  on a preview API and is deliberately not used here. */
export async function markAsJunk(accessToken: string, messageId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/markAsJunk`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ moveToJunk: true }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "report the message as junk");
}

/** Report a message onward to Microsoft's filters: junk, phishing, or a
 *  false positive rescued from Junk.
 *
 *  This one lives on /beta, which Microsoft says outright is not supported
 *  for production use and may change. It is here because there is no v1.0
 *  equivalent: markAsJunk covers only the junk half, and its notJunk twin was
 *  deprecated with a cutoff that has already passed. The caller treats a
 *  failure as a real answer rather than retrying elsewhere, and the UI says
 *  which of these is the preview one. The action value is `phish`, not
 *  `phishing`. It needs Mail.ReadWrite, which is already granted. */
export async function reportMessageBeta(accessToken: string, messageId: string, action: "junk" | "notJunk" | "phish", moveOut: boolean): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/beta/me/messages/${encodeURIComponent(messageId)}/reportMessage`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ IsMessageMoveRequested: moveOut, ReportAction: action }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, action === "phish" ? "report the message as phishing" : action === "notJunk" ? "report the message as not junk" : "report the message as junk");
}

/** Someone's profile picture at 48 pixels, as a data url, or null when they
 *  have none.
 *
 *  The fixed-size endpoint rather than the default one, which hands back the
 *  largest photo uploaded: a 648-pixel portrait behind a 32-pixel circle is
 *  a hundred times the bytes for no visible gain. A 404 means this person
 *  simply has no photo, which is the common case for anyone outside the
 *  organization and is not an error. A 403 means the permission was never
 *  granted, and that is worth telling the caller so it can stop asking. */
export async function getUserPhoto(accessToken: string, address: string): Promise<string | null> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(address)}/photos/48x48/$value`,
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status === 403 || r.status === 401) throw new Error("no-photo-permission");
	if (r.status >= 400) return null;
	try {
		return `data:image/jpeg;base64,${arrayBufferToBase64(r.arrayBuffer)}`;
	} catch {
		return null;
	}
}

/* ---------- automatic replies ---------- */

export interface AutoReply {
	status: "disabled" | "alwaysEnabled" | "scheduled";
	externalAudience: "none" | "contactsOnly" | "all";
	internalReplyMessage: string;
	externalReplyMessage: string;
	scheduledStartDateTime?: { dateTime: string; timeZone: string };
	scheduledEndDateTime?: { dateTime: string; timeZone: string };
}

/** The mailbox's out-of-office setting. Reads and writes under
 *  MailboxSettings, the same permission inbox rules needed. */
export async function getAutoReply(accessToken: string): Promise<AutoReply | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailboxSettings/automaticRepliesSetting",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "read the automatic replies");
	const v = bodyJson(r);
	if (!v) return null;
	return {
		status: (v.status as AutoReply["status"]) ?? "disabled",
		externalAudience: (v.externalAudience as AutoReply["externalAudience"]) ?? "none",
		internalReplyMessage: String(v.internalReplyMessage ?? ""),
		externalReplyMessage: String(v.externalReplyMessage ?? ""),
		scheduledStartDateTime: v.scheduledStartDateTime as AutoReply["scheduledStartDateTime"],
		scheduledEndDateTime: v.scheduledEndDateTime as AutoReply["scheduledEndDateTime"],
	};
}

export async function setAutoReply(accessToken: string, setting: AutoReply): Promise<void> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailboxSettings",
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ automaticRepliesSetting: setting }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "set the automatic replies");
}

/** The mailbox's saved contacts, flattened to one entry per address, since
 *  the address is what you write to. A contact with no address at all is not
 *  useful here and is dropped. */
export async function listContacts(accessToken: string): Promise<{ name: string; email: string; company?: string; title?: string; phone?: string }[]> {
	const raw = await pagedGet(
		"https://graph.microsoft.com/v1.0/me/contacts?$top=100&$select=displayName,emailAddresses,companyName,jobTitle,businessPhones,mobilePhone",
		accessToken,
		{},
		20
	);
	const out: { name: string; email: string; company?: string; title?: string; phone?: string }[] = [];
	for (const c of raw) {
		const o = c as {
			displayName?: string;
			emailAddresses?: { address?: string; name?: string }[];
			companyName?: string;
			jobTitle?: string;
			businessPhones?: string[];
			mobilePhone?: string;
		};
		const phone = o.mobilePhone || o.businessPhones?.[0];
		for (const e of o.emailAddresses ?? []) {
			const email = (e.address ?? "").trim();
			if (!email.includes("@")) continue;
			out.push({
				name: (o.displayName || e.name || email).trim(),
				email,
				company: o.companyName?.trim() || undefined,
				title: o.jobTitle?.trim() || undefined,
				phone: phone?.trim() || undefined,
			});
		}
	}
	return out;
}

/** The unread messages in a folder, ids only.
 *
 *  Capped, because "mark all read" on a folder holding twenty thousand is a
 *  request per twenty of them, and a job that long should be declined rather
 *  than started and abandoned halfway. */
export async function listUnreadIdsInFolder(accessToken: string, folderId: string, cap = 2000): Promise<{ ids: string[]; complete: boolean }> {
	let url: string | null =
		`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages?$filter=isRead%20eq%20false&$select=id&$top=100`;
	const ids: string[] = [];
	for (let page = 0; url && page < 40; page++) {
		const r: { status: number; json: unknown } = await requestUrl({ url, headers: { Authorization: `Bearer ${accessToken}` }, throw: false });
		if (r.status >= 400) throw writeError(r as { status: number; json: unknown }, "read the folder");
		const j = bodyJson(r);
		for (const v of (j?.value as { id?: string }[] | undefined) ?? []) if (v.id) ids.push(v.id);
		if (ids.length >= cap) return { ids: ids.slice(0, cap), complete: false };
		url = (j?.["@odata.nextLink"] as string | undefined) ?? null;
	}
	return { ids, complete: true };
}

/** Mark up to twenty messages read or unread in one round trip. */
export async function patchReadBatch(accessToken: string, ids: readonly string[], read: boolean): Promise<{ ok: string[]; failed: string[] }> {
	if (!ids.length) return { ok: [], failed: [] };
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/$batch",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(buildReadPatchBatch(ids, read)),
		throw: false,
	});
	if (r.status >= 400) return { ok: [], failed: [...ids] };
	return parseWriteBatch(r.json, ids);
}

/* ---------- Outlook notes (sticky notes) ---------- */

/** Outlook's notes have no endpoint of their own and never did, which is why
 *  everyone concludes Graph cannot reach them. They are ordinary items in the
 *  Notes folder wearing the message class IPM.StickyNote, so the mail API
 *  reaches them perfectly well under the permission mail already has. */
export interface StickyNote {
	id: string;
	title: string;
	preview: string;
	body: string;
	changedMs: number;
}

const NOTE_CLASS = "IPM.StickyNote";
/** PidTagMessageClass. What makes an item a note rather than a message. */
const MESSAGE_CLASS_PROP = "String 0x001A";

/** The mailbox's Notes folder, by name. There is no well-known id for it. */
export async function findNotesFolder(accessToken: string): Promise<string | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailFolders?$top=100&$select=id,displayName",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "find the Notes folder");
	const list = (bodyJson(r)?.value as { id?: string; displayName?: string }[] | undefined) ?? [];
	return list.find((f) => (f.displayName ?? "").toLowerCase() === "notes")?.id ?? null;
}

export async function listStickyNotes(accessToken: string, folderId: string): Promise<StickyNote[]> {
	const raw = await pagedGet(
		`https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages?$top=100&$select=id,subject,bodyPreview,lastModifiedDateTime&$orderby=${encodeURIComponent("lastModifiedDateTime desc")}`,
		accessToken,
		{},
		10
	);
	return raw
		.map((v) => {
			const o = v as { id?: string; subject?: string; bodyPreview?: string; lastModifiedDateTime?: string };
			const ms = Date.parse(o.lastModifiedDateTime ?? "");
			return {
				id: String(o.id ?? ""),
				title: (o.subject ?? "").trim() || (o.bodyPreview ?? "").split("\n")[0]?.trim() || "(empty note)",
				preview: (o.bodyPreview ?? "").replace(/\s+/g, " ").trim(),
				body: "",
				changedMs: Number.isFinite(ms) ? ms : 0,
			};
		})
		.filter((n) => n.id);
}

export async function getStickyNoteBody(accessToken: string, id: string): Promise<string> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(id)}?$select=body`,
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "read the note");
	const body = bodyJson(r)?.body as { content?: string; contentType?: string } | undefined;
	return body?.content ?? "";
}

/** Make a note that Outlook will recognize as one.
 *
 *  A plain POST here would create a draft message sitting in the Notes
 *  folder, which Outlook shows as mail rather than as a note. Setting the
 *  message class is the whole difference. */
export async function createStickyNote(accessToken: string, folderId: string, text: string): Promise<void> {
	const firstLine = text.split("\n")[0]?.trim().slice(0, 120) || "Note";
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders/${encodeURIComponent(folderId)}/messages`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			subject: firstLine,
			body: { contentType: "text", content: text },
			singleValueExtendedProperties: [{ id: MESSAGE_CLASS_PROP, value: NOTE_CLASS }],
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "create the note");
}

/* ---------- tasks (Microsoft To Do) ---------- */

export interface TodoList {
	id: string;
	name: string;
	isDefault: boolean;
}

export interface TodoTask {
	id: string;
	listId: string;
	title: string;
	done: boolean;
	/** Local midnight of the due day, or null. Graph carries a
	 *  dateTimeTimeZone here and the time half is meaningless for a due
	 *  date, so only the day is kept. */
	dueMs: number | null;
	importance: "low" | "normal" | "high";
	note: string;
}

export async function listTodoLists(accessToken: string): Promise<TodoList[]> {
	const raw = await pagedGet("https://graph.microsoft.com/v1.0/me/todo/lists?$top=100", accessToken);
	return raw
		.map((v) => {
			const o = v as { id?: string; displayName?: string; wellknownListName?: string };
			return { id: String(o.id ?? ""), name: String(o.displayName ?? "Tasks"), isDefault: o.wellknownListName === "defaultList" };
		})
		.filter((l) => l.id);
}

const todoFrom = (v: Record<string, unknown>, listId: string): TodoTask => {
	const due = v.dueDateTime as { dateTime?: string } | undefined;
	// Graph returns the due date at midnight UTC; only the day matters, so it
	// is read as a day rather than an instant, or a task due Friday shows as
	// Thursday for anyone west of Greenwich
	const day = (due?.dateTime ?? "").slice(0, 10);
	const body = v.body as { content?: string } | undefined;
	return {
		id: String(v.id ?? ""),
		listId,
		title: String(v.title ?? "(untitled)"),
		done: v.status === "completed",
		dueMs: /^\d{4}-\d{2}-\d{2}$/.test(day) ? msOfKey(day) : null,
		importance: v.importance === "high" || v.importance === "low" ? v.importance : "normal",
		note: String(body?.content ?? ""),
	};
};

export async function listTodoTasks(accessToken: string, listId: string): Promise<TodoTask[]> {
	const raw = await pagedGet(`https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks?$top=100`, accessToken, {}, 10);
	return raw.map((v) => todoFrom(v as Record<string, unknown>, listId)).filter((t) => t.id);
}

export async function createTodoTask(accessToken: string, listId: string, title: string, dueMs: number | null): Promise<void> {
	const body: Record<string, unknown> = { title };
	if (dueMs != null) body.dueDateTime = { dateTime: `${keyOfMs(dueMs)}T00:00:00.0000000`, timeZone: "UTC" };
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "add the task");
}

export async function setTodoTaskDone(accessToken: string, listId: string, taskId: string, done: boolean): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ status: done ? "completed" : "notStarted" }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "update the task");
}

export async function deleteTodoTask(accessToken: string, listId: string, taskId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the task");
}

/* ---------- categories ---------- */

export interface OutlookCategory {
	id: string;
	displayName: string;
	/** Graph names colors "preset0" through "preset24" rather than giving a
	 *  value, so the palette has to be carried on this side. */
	color: string;
}

/** The mailbox's own category list, the same one Outlook's Categorize menu
 *  shows. Reads under MailboxSettings, which inbox rules already granted. */
export async function listMasterCategories(accessToken: string): Promise<OutlookCategory[]> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "read the categories");
	const list = (bodyJson(r)?.value as Record<string, unknown>[] | undefined) ?? [];
	return list.map((v) => ({ id: String(v.id ?? ""), displayName: String(v.displayName ?? ""), color: String(v.color ?? "") })).filter((c) => c.displayName);
}

/** Every message carrying a category, with its current category list.
 *
 *  `categories/any(a:a eq '...')` is the supported filter for this. Paged to
 *  a cap so a mailbox with a category on twenty thousand messages reports a
 *  refusal rather than grinding: a replace that silently covered only part
 *  of the mailbox would be worse than one that declined. */
export async function findMessagesByCategory(accessToken: string, name: string, cap = 2000): Promise<{ hits: { id: string; categories: string[] }[]; complete: boolean }> {
	const filter = encodeURIComponent(`categories/any(a:a eq '${name.replace(/'/g, "''")}')`);
	let url: string | null = `https://graph.microsoft.com/v1.0/me/messages?$filter=${filter}&$select=id,categories&$top=100`;
	const hits: { id: string; categories: string[] }[] = [];
	for (let page = 0; url && page < 40; page++) {
		const r: { status: number; json: unknown } = await requestUrl({ url, headers: { Authorization: `Bearer ${accessToken}` }, throw: false });
		if (r.status >= 400) throw writeError(r as { status: number; json: unknown }, "find the messages in that category");
		const j = bodyJson(r);
		for (const v of (j?.value as { id?: string; categories?: string[] }[] | undefined) ?? []) {
			if (v.id) hits.push({ id: v.id, categories: v.categories ?? [] });
		}
		if (hits.length >= cap) return { hits: hits.slice(0, cap), complete: false };
		url = (j?.["@odata.nextLink"] as string | undefined) ?? null;
	}
	return { hits, complete: true };
}

/** Rewrite the categories of up to twenty messages in one round trip. */
export async function patchCategoriesBatch(accessToken: string, items: readonly { id: string; categories: string[] }[]): Promise<{ ok: string[]; failed: string[] }> {
	if (!items.length) return { ok: [], failed: [] };
	const ids = items.map((i) => i.id);
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/$batch",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(buildCategoryPatchBatch(items)),
		throw: false,
	});
	if (r.status >= 400) return { ok: [], failed: [...ids] };
	return parseWriteBatch(r.json, ids);
}

export async function createCategory(accessToken: string, displayName: string, color: string): Promise<OutlookCategory | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ displayName, color }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "create the category");
	const v = bodyJson(r);
	return v ? { id: String(v.id ?? ""), displayName: String(v.displayName ?? ""), color: String(v.color ?? "") } : null;
}

/** Recolor a category. Color is the only writable property: Graph refuses to
 *  change a category's name once it exists, which is why this plugin offers
 *  no rename rather than offering one that quietly fails. */
export async function updateCategoryColor(accessToken: string, categoryId: string, color: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/outlook/masterCategories/${encodeURIComponent(categoryId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ color }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "recolor the category");
}

export async function deleteCategory(accessToken: string, categoryId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/outlook/masterCategories/${encodeURIComponent(categoryId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the category");
}

/** Set a message's categories. The whole list is written every time, since
 *  that is what the property is: PATCHing one on means sending the others
 *  back with it. */
export async function setMessageCategories(accessToken: string, messageId: string, categories: string[]): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ categories }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "set the categories");
}

/* ---------- inbox rules ---------- */

/** A rule as the mailbox stores it. `conditions`, `actions`, and
 *  `exceptions` stay as loose records on purpose: Outlook can set far more in
 *  them than this plugin offers, and anything it does not understand has to
 *  survive a round trip through here untouched. */
export interface GraphRule {
	id: string;
	displayName: string;
	sequence: number;
	isEnabled: boolean;
	conditions: Record<string, unknown>;
	actions: Record<string, unknown>;
	exceptions: Record<string, unknown>;
}

const ruleFrom = (v: Record<string, unknown>): GraphRule => ({
	id: String(v.id ?? ""),
	displayName: String(v.displayName ?? "(unnamed rule)"),
	sequence: Number(v.sequence ?? 0),
	isEnabled: v.isEnabled !== false,
	conditions: (v.conditions as Record<string, unknown>) ?? {},
	actions: (v.actions as Record<string, unknown>) ?? {},
	exceptions: (v.exceptions as Record<string, unknown>) ?? {},
});

/** The inbox's rules, in the order the mailbox runs them. */
export async function listMessageRules(accessToken: string): Promise<GraphRule[]> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "read the rules");
	const list = (bodyJson(r)?.value as Record<string, unknown>[] | undefined) ?? [];
	return list.map(ruleFrom).sort((a, b) => a.sequence - b.sequence);
}

export async function createMessageRule(accessToken: string, body: Record<string, unknown>): Promise<GraphRule | null> {
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "create the rule");
	const j = bodyJson(r);
	return j ? ruleFrom(j) : null;
}

export async function updateMessageRule(accessToken: string, ruleId: string, body: Record<string, unknown>): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "update the rule");
}

export async function deleteMessageRule(accessToken: string, ruleId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules/${encodeURIComponent(ruleId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the rule");
}

/** Into Deleted Items, recoverable there; exactly Outlook's Delete. */
export async function deleteMessage(accessToken: string, messageId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/move`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ destinationId: "deleteditems" }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the message");
}

/** Gone for real: past Deleted Items, unrecoverable from the mailbox. */
export async function permanentDeleteMessage(accessToken: string, messageId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/permanentDelete`,
		method: "POST",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "permanently delete the message");
}

/** Reply to everyone on the thread; Graph handles recipients and quoting. */
export async function replyAllMessage(accessToken: string, messageId: string, commentHtml: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/replyAll`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ comment: commentHtml }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "send the reply");
}

/** Forward with a comment; Graph carries the original body and attachments. */
export async function forwardMessage(accessToken: string, messageId: string, commentHtml: string, to: string[]): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/forward`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ comment: commentHtml, toRecipients: to.map((a) => ({ emailAddress: { address: a } })) }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "forward the message");
}

/** Reply to the sender in-thread; Graph handles quoting and the subject. */
export async function replyMessage(accessToken: string, messageId: string, commentHtml: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/reply`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ comment: commentHtml }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "send the reply");
}

/** Send a new message as the signed-in user, filed in their own Sent Items.
 *  Graph answers 202 and queues: accepted for delivery, not delivered. */
/** Receipts a message can ask for.
 *
 *  These are the standard mail fields, not a tracking pixel: they travel as
 *  headers, the recipient's own client decides whether to honor them, and
 *  every serious one asks the person first and lets them refuse. Which is
 *  also why a receipt never arriving means nothing in particular. */
export interface Receipts {
	read?: boolean;
	delivery?: boolean;
}

const receiptFields = (r?: Receipts) => ({
	...(r?.read ? { isReadReceiptRequested: true } : {}),
	...(r?.delivery ? { isDeliveryReceiptRequested: true } : {}),
});

export async function sendGraphMail(accessToken: string, m: { to: string[]; cc?: string[]; bcc?: string[]; subject: string; html: string; receipts?: Receipts }): Promise<void> {
	const box = (a: string) => ({ emailAddress: { address: a } });
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/sendMail",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			message: {
				subject: m.subject,
				body: { contentType: "HTML", content: m.html },
				toRecipients: m.to.map(box),
				...(m.cc?.length ? { ccRecipients: m.cc.map(box) } : {}),
				...(m.bcc?.length ? { bccRecipients: m.bcc.map(box) } : {}),
				...receiptFields(m.receipts),
			},
			saveToSentItems: true,
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "send the mail");
}

export interface DraftMessage {
	id: string;
	subject: string;
	bodyHtml: string;
	to: string[];
	cc: string[];
}

const draftFrom = (v: Record<string, unknown>): DraftMessage => {
	const addrs = (list: unknown): string[] => (((list as { emailAddress?: { address?: string } }[] | undefined) ?? []).map((r) => r.emailAddress?.address ?? "").filter(Boolean));
	return {
		id: String(v.id ?? ""),
		subject: String(v.subject ?? ""),
		bodyHtml: String((v.body as { content?: string } | undefined)?.content ?? ""),
		to: addrs(v.toRecipients),
		cc: addrs(v.ccRecipients),
	};
};

/** A reply, reply-all, or forward draft: Graph prefills the recipients, the
 *  RE/FW subject, and the quoted original, exactly what Outlook edits. */
export async function createDraftReply(accessToken: string, messageId: string, kind: "reply" | "replyAll" | "forward"): Promise<DraftMessage> {
	const verb = kind === "reply" ? "createReply" : kind === "replyAll" ? "createReplyAll" : "createForward";
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/${verb}`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: "{}",
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "open the reply");
	return draftFrom((r.json as Record<string, unknown>) ?? {});
}

export async function updateDraft(accessToken: string, draftId: string, patch: { subject: string; html: string; to: string[]; cc: string[]; bcc?: string[]; receipts?: Receipts }): Promise<void> {
	const box = (a: string) => ({ emailAddress: { address: a } });
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			subject: patch.subject,
			body: { contentType: "HTML", content: patch.html },
			toRecipients: patch.to.map(box),
			ccRecipients: patch.cc.map(box),
			bccRecipients: (patch.bcc ?? []).map(box),
			// written every time, not only when asked for: a draft that once
			// requested a receipt has to be able to stop
			isReadReceiptRequested: !!patch.receipts?.read,
			isDeliveryReceiptRequested: !!patch.receipts?.delivery,
		}),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "save the draft");
}

export async function sendDraft(accessToken: string, draftId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}/send`,
		method: "POST",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "send the mail");
}

export async function deleteDraft(accessToken: string, draftId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "discard the draft");
}

export interface ScheduleInfo {
	email: string;
	availabilityView: string;
	/** Set when the tenant offers no visibility into this mailbox. */
	error: string | null;
}

/** Free/busy for colleagues over a window (the scheduling assistant). Works
 *  for mailboxes the tenant lets you see; outsiders come back as errors, not
 *  failures. */
export async function getSchedule(accessToken: string, emails: string[], startMs: number, endMs: number, intervalMin = 30): Promise<ScheduleInfo[]> {
	const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/calendar/getSchedule",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}`, Prefer: `outlook.timezone="${tz}"` },
		body: JSON.stringify({
			schedules: emails,
			startTime: { dateTime: wallOfMs(startMs), timeZone: tz },
			endTime: { dateTime: wallOfMs(endMs), timeZone: tz },
			availabilityViewInterval: intervalMin,
		}),
		throw: false,
	});
	if (r.status >= 400) {
		const j = bodyJson(r);
		const err = j?.error as { code?: string; message?: string } | undefined;
		throw new GraphError(err?.message || `Could not read availability (${r.status}).`, err?.code);
	}
	const value = (bodyJson(r)?.value as Record<string, unknown>[] | undefined) ?? [];
	return emails.map((email, i) => {
		const v = value[i] ?? {};
		const err = v.error as { message?: string } | undefined;
		const view = typeof v.availabilityView === "string" ? v.availabilityView : "";
		return { email, availabilityView: view, error: err?.message ?? (view ? null : "no visibility") };
	});
}

/** One event by id in the device's zone; how series-wide edits load the
 *  master before opening the editor. */
export async function getEvent(accessToken: string, eventId: string): Promise<unknown> {
	const tz = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(eventId)}?$select=${EVENT_SELECT}`,
		method: "GET",
		headers: { Authorization: `Bearer ${accessToken}`, Prefer: `outlook.timezone="${tz}"` },
		throw: false,
	});
	if (r.status >= 400) {
		const j = bodyJson(r);
		const err = j?.error as { code?: string; message?: string } | undefined;
		throw new GraphError(err?.message || `Could not read the event (${r.status}).`, err?.code);
	}
	return r.json as unknown;
}

/* ---------------- writes ---------------- */

function writeError(r: { status: number; json: unknown }, doing: string): GraphError {
	const j = bodyJson(r);
	const err = j?.error as { code?: string; message?: string } | undefined;
	// 403 is almost always the permission, not the sign-in: the connection was
	// made under the read-only scope, or the event belongs to someone else
	if (r.status === 403) return new GraphError(err?.message || `Microsoft 365 refused to ${doing} (403). Reconnect to grant edit access, or the event may not be yours to change.`, err?.code);
	return new GraphError(err?.message || `Could not ${doing} (${r.status}).`, err?.code);
}

/** Create an event; null calendarId targets the default calendar. */
export async function createEvent(accessToken: string, calendarId: string | null, body: Record<string, unknown>): Promise<void> {
	const url = calendarId ? `https://graph.microsoft.com/v1.0/me/calendars/${encodeURIComponent(calendarId)}/events` : "https://graph.microsoft.com/v1.0/me/events";
	const r = await requestUrl({
		url,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "create the event");
}

/** Patch an event (or one occurrence of a series, via the instance id from
 *  calendarView, which Graph turns into an exception). */
export async function updateEvent(accessToken: string, eventId: string, body: Record<string, unknown>): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(eventId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify(body),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "update the event");
}

/** Respond to an invite (works on an instance id for one occurrence). The
 *  organizer is told, exactly as responding in Outlook would. */
export async function respondEvent(accessToken: string, eventId: string, action: "accept" | "tentativelyAccept" | "decline"): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(eventId)}/${action}`,
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ sendResponse: true }),
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "send the response");
}

/** Delete an event, or cancel one occurrence when given an instance id. Lands
 *  in Deleted Items on the server, so it is recoverable there. */
export async function deleteEvent(accessToken: string, eventId: string): Promise<void> {
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/events/${encodeURIComponent(eventId)}`,
		method: "DELETE",
		headers: { Authorization: `Bearer ${accessToken}` },
		throw: false,
	});
	if (r.status >= 400) throw writeError(r, "delete the event");
}
