/* CalDAV XML: request bodies and multistatus parsing. Pure: no Obsidian
 * imports, unit-tested with Node.
 *
 * Parsing is a tolerant, namespace-agnostic regex layer rather than a real
 * XML parser: DOMParser does not exist in the Node test runner, DAV responses
 * are machine-written and small, and servers disagree only about namespace
 * PREFIXES (d:, D:, none), which local-name matching sidesteps entirely. */

export function xmlEscape(s: string): string {
	return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function xmlUnescape(s: string): string {
	return s
		.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
		.replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
		.replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, "&");
}

const NS = "(?:[A-Za-z0-9_-]+:)?";

/** Inner text of the first <local> element, any namespace prefix; null if absent. */
export function firstTag(xml: string, local: string): string | null {
	const m = xml.match(new RegExp(`<${NS}${local}(?:\\s[^>]*)?>([\\s\\S]*?)</${NS}${local}>`, "i"));
	return m ? m[1] : null;
}

/** Inner texts of every <local> element, any namespace prefix. */
export function allTags(xml: string, local: string): string[] {
	const out: string[] = [];
	const re = new RegExp(`<${NS}${local}(?:\\s[^>]*)?>([\\s\\S]*?)</${NS}${local}>`, "gi");
	let m: RegExpExecArray | null;
	while ((m = re.exec(xml))) out.push(m[1]);
	return out;
}

/** Whether a self-closed or empty <local/> element is present (resourcetype
 *  markers like <cal:calendar/>). */
export function hasEmptyTag(xml: string, local: string): boolean {
	return new RegExp(`<${NS}${local}(?:\\s[^>]*)?/>|<${NS}${local}(?:\\s[^>]*)?></${NS}${local}>`, "i").test(xml);
}

/** The inner body of each <response> in a multistatus document. */
export function responses(xml: string): string[] {
	return allTags(xml, "response");
}

/* ---------- request bodies ---------- */

export const PROPFIND_PRINCIPAL = `<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:"><d:prop><d:current-user-principal/></d:prop></d:propfind>`;

export const PROPFIND_HOME = `<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav"><d:prop><c:calendar-home-set/></d:prop></d:propfind>`;

export const PROPFIND_CALENDARS = `<?xml version="1.0" encoding="utf-8"?>
<d:propfind xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav" xmlns:a="http://apple.com/ns/ical/">
<d:prop><d:displayname/><d:resourcetype/><a:calendar-color/><c:supported-calendar-component-set/></d:prop>
</d:propfind>`;

/** "YYYYMMDDTHHMMSSZ" for time-range filters. */
export function toBasicUtc(ms: number): string {
	const d = new Date(ms);
	const p = (n: number) => String(n).padStart(2, "0");
	return `${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}T${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}

/** A calendar-query REPORT for VEVENTs intersecting [startMs, endMs). The
 *  server returns whole series when any instance intersects; expansion stays
 *  client-side so overrides survive. */
export function buildCalendarQuery(startMs: number, endMs: number): string {
	return `<?xml version="1.0" encoding="utf-8"?>
<c:calendar-query xmlns:d="DAV:" xmlns:c="urn:ietf:params:xml:ns:caldav">
<d:prop><d:getetag/><c:calendar-data/></d:prop>
<c:filter><c:comp-filter name="VCALENDAR"><c:comp-filter name="VEVENT">
<c:time-range start="${toBasicUtc(startMs)}" end="${toBasicUtc(endMs)}"/>
</c:comp-filter></c:comp-filter></c:filter>
</c:calendar-query>`;
}

/* ---------- response parsing ---------- */

/** Resolve a DAV href against the request URL: absolute stays, root-relative
 *  takes the origin, anything else appends to the base directory. */
export function resolveHref(base: string, href: string): string {
	const h = xmlUnescape(href.trim());
	if (/^https?:\/\//i.test(h)) return h;
	const om = base.match(/^(https?:\/\/[^/]+)/i);
	const origin = om ? om[1] : base.replace(/\/+$/, "");
	if (h.startsWith("/")) return origin + h;
	return base.replace(/\/+$/, "") + "/" + h;
}

export interface DavCollection {
	/** Absolute URL of the calendar collection. */
	href: string;
	name: string;
	/** #rrggbb when the server offers one. */
	color: string | null;
}

/** Calendar collections in a PROPFIND Depth:1 multistatus: resourcetype must
 *  carry <calendar/>, and a supported-calendar-component-set that exists but
 *  lacks VEVENT excludes the collection (task-only and journal collections). */
export function parseCollections(xml: string, baseUrl: string): DavCollection[] {
	const out: DavCollection[] = [];
	for (const body of responses(xml)) {
		const href = firstTag(body, "href");
		if (!href) continue;
		const rtype = firstTag(body, "resourcetype") ?? "";
		if (!hasEmptyTag(rtype, "calendar")) continue;
		const compSet = firstTag(body, "supported-calendar-component-set");
		if (compSet != null && !/name="VEVENT"/i.test(compSet)) continue;
		const abs = resolveHref(baseUrl, href);
		const rawName = firstTag(body, "displayname");
		const name = rawName?.trim() ? xmlUnescape(rawName.trim()) : decodeURIComponent(abs.replace(/\/+$/, "").split("/").pop() ?? "Calendar");
		const rawColor = firstTag(body, "calendar-color")?.trim() ?? "";
		const colorMatch = rawColor.match(/^#[0-9a-fA-F]{6}/);
		out.push({ href: abs, name, color: colorMatch ? colorMatch[0].toLowerCase() : null });
	}
	return out;
}

/** Every calendar-data payload in a REPORT multistatus, XML-unescaped back to
 *  raw iCalendar text. */
export function parseCalendarData(xml: string): string[] {
	return allTags(xml, "calendar-data")
		.map((t) => xmlUnescape(t).trim())
		.filter((t) => t.includes("BEGIN:VCALENDAR"));
}

/** The principal href out of a current-user-principal PROPFIND; null when the
 *  server does not say. */
export function parsePrincipal(xml: string, baseUrl: string): string | null {
	const block = firstTag(xml, "current-user-principal");
	const href = block ? firstTag(block, "href") : null;
	return href ? resolveHref(baseUrl, href) : null;
}

/** The calendar-home-set href out of its PROPFIND; null when absent. */
export function parseHomeSet(xml: string, baseUrl: string): string | null {
	const block = firstTag(xml, "calendar-home-set");
	const href = block ? firstTag(block, "href") : null;
	return href ? resolveHref(baseUrl, href) : null;
}
