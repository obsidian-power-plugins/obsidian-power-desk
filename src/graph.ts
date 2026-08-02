import { requestUrl } from "obsidian";
import { buildMessageBatch, parseMessageBatch } from "./core";
import type { BatchOutcome } from "./core";
import { wallOfMs } from "./core";

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
export const GRAPH_SCOPE = "offline_access openid profile Calendars.ReadWrite Mail.ReadWrite Mail.Send";
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
	"id,subject,from,receivedDateTime,bodyPreview,isRead,webLink,hasAttachments,parentFolderId,importance,flag,toRecipients,conversationId,inferenceClassification";

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
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}?$select=subject,from,toRecipients,receivedDateTime,body,webLink`,
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
export async function sendGraphMail(accessToken: string, m: { to: string[]; cc?: string[]; subject: string; html: string }): Promise<void> {
	const box = (a: string) => ({ emailAddress: { address: a } });
	const r = await requestUrl({
		url: "https://graph.microsoft.com/v1.0/me/sendMail",
		method: "POST",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({
			message: { subject: m.subject, body: { contentType: "HTML", content: m.html }, toRecipients: m.to.map(box), ...(m.cc?.length ? { ccRecipients: m.cc.map(box) } : {}) },
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

export async function updateDraft(accessToken: string, draftId: string, patch: { subject: string; html: string; to: string[]; cc: string[] }): Promise<void> {
	const box = (a: string) => ({ emailAddress: { address: a } });
	const r = await requestUrl({
		url: `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(draftId)}`,
		method: "PATCH",
		contentType: "application/json",
		headers: { Authorization: `Bearer ${accessToken}` },
		body: JSON.stringify({ subject: patch.subject, body: { contentType: "HTML", content: patch.html }, toRecipients: patch.to.map(box), ccRecipients: patch.cc.map(box) }),
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
