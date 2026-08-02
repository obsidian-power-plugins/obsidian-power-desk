/* Power Calendar: pure view logic. No Obsidian imports, everything here is
 * unit-tested with Node (npm test). Anything that touches the Obsidian API
 * lives in main.ts; anything that talks HTTP lives in graph.ts / caldav.ts. */

/** One concrete event instance, whatever its source. Recurring series arrive
 *  already expanded: each visible occurrence is its own PCEvent. */
export interface PCEvent {
	/** Stable per instance: provider id, or uid plus occurrence start. */
	id: string;
	sourceId: string;
	/** The calendar's user-facing name (account or feed), for the detail card. */
	calendarName?: string;
	/** Resolved chip color (hex). */
	color?: string;
	title: string;
	allDay: boolean;
	/** Local epoch ms. All-day events run local midnight to local midnight. */
	startMs: number;
	/** Exclusive: the first instant no longer inside the event. */
	endMs: number;
	location?: string;
	organizer?: string;
	attendees?: string[];
	description?: string;
	/** Provider page for the event (Outlook web link, ICS URL property). */
	url?: string;
	/** Teams / Meet / Zoom / Webex link when one is found. */
	joinUrl?: string;
	recurring?: boolean;
	/** The id of the series master this instance expands from, when known;
	 *  patching or deleting IT edits the whole series. */
	seriesId?: string;
	/** Outlook categories, exactly as the mailbox names them. */
	categories?: string[];
	tentative?: boolean;
	/** The signed-in user declined this invite (Graph only knows this). */
	declined?: boolean;
	/** This event can be edited: a writable source AND the user organizes it. */
	canEdit?: boolean;
	/** Attendees with their addresses, for edit prefills and RSVP writes. */
	attendeeDetail?: { name?: string; email?: string }[];
	/** The signed-in user's own response to this invite. */
	myResponse?: "accepted" | "tentative" | "declined" | "none";
	/** An invite the signed-in user can respond to (not their own meeting). */
	canRsvp?: boolean;
	/** Marked free / transparent: it does not block availability. */
	transparent?: boolean;
	/** Set when the event IS a vault note (a vault-notes source): the path to
	 *  open, and the sign that provider actions do not apply. */
	notePath?: string;
}

/* ---------- day-key algebra (UTC-based over YYYY-MM-DD keys, so DST can
 * never skew a span; keys themselves always mean LOCAL days) ---------- */

/** YYYY-MM-DD to whole days since the epoch. */
export function dayNum(key: string): number {
	const y = +key.slice(0, 4);
	const m = +key.slice(5, 7) - 1;
	const d = +key.slice(8, 10);
	return Math.round(Date.UTC(y, m, d) / 86400000);
}

/** Whole days since the epoch back to YYYY-MM-DD. */
export function keyOfDayNum(n: number): string {
	const d = new Date(n * 86400000);
	return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
}

export function addDays(key: string, days: number): string {
	return keyOfDayNum(dayNum(key) + days);
}

/** Days from a to b (positive when b is later). */
export function dayDiff(a: string, b: string): number {
	return dayNum(b) - dayNum(a);
}

/** 0 = Sunday ... 6 = Saturday. */
export function dayOfWeek(key: string): number {
	return (dayNum(key) + 4) % 7;
}

/** ISO 8601 week number: the week that contains the Thursday. */
export function isoWeekNum(key: string): number {
	const n = dayNum(key);
	const dowMon = (n + 3) % 7; // 0 = Monday
	const thursday = n - dowMon + 3;
	const jan1 = dayNum(`${keyOfDayNum(thursday).slice(0, 4)}-01-01`);
	return Math.floor((thursday - jan1) / 7) + 1;
}

/** The seven day-keys of the week containing `key`, starting Monday or Sunday. */
export function weekDays(key: string, mondayStart: boolean): string[] {
	const dow = dayOfWeek(key); // 0 = Sunday
	const offset = mondayStart ? (dow + 6) % 7 : dow;
	const first = addDays(key, -offset);
	return Array.from({ length: 7 }, (_, i) => addDays(first, i));
}

export interface DayCell {
	/** YYYY-MM-DD */
	key: string;
	day: number;
	inMonth: boolean;
}

/** A fixed 6x7 month grid (42 cells) starting on the configured weekday. */
export function monthGrid(year: number, month0: number, weekStartsMonday: boolean): DayCell[] {
	const first = new Date(year, month0, 1);
	let lead = first.getDay() - (weekStartsMonday ? 1 : 0);
	if (lead < 0) lead += 7;
	const cells: DayCell[] = [];
	for (let i = 0; i < 42; i++) {
		const d = new Date(year, month0, 1 - lead + i);
		const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
		cells.push({ key, day: d.getDate(), inMonth: d.getMonth() === month0 });
	}
	return cells;
}

/* ---------- local-time bridges: epochs and keys meet here, and only here ---------- */

/** First YYYY-MM-DD found in a string (a daily note's filename); null if none. */
export function dateKeyOf(raw: string): string | null {
	const m = raw.match(/(\d{4})-(\d{2})-(\d{2})/);
	if (!m) return null;
	const mo = +m[2];
	const d = +m[3];
	if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
	return `${m[1]}-${m[2]}-${m[3]}`;
}

/** The LOCAL day a Date falls on. */
export function keyOfDate(d: Date): string {
	return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

export function keyOfMs(ms: number): string {
	return keyOfDate(new Date(ms));
}

/** Local midnight of a day key, as epoch ms. */
export function msOfKey(key: string): number {
	return new Date(+key.slice(0, 4), +key.slice(5, 7) - 1, +key.slice(8, 10)).getTime();
}

/** Local minutes since midnight of an instant. */
export function minutesOfMs(ms: number): number {
	const d = new Date(ms);
	return d.getHours() * 60 + d.getMinutes();
}

/* ---------- event geometry ---------- */

/** The local days an event touches, endKey INCLUSIVE. An event ending exactly
 *  at midnight belongs to the day it ended on, not the day it merely grazed. */
export function eventDaySpan(ev: { startMs: number; endMs: number }): { startKey: string; endKey: string } {
	const startKey = keyOfMs(ev.startMs);
	const endKey = keyOfMs(Math.max(ev.startMs, ev.endMs - 1));
	return { startKey, endKey };
}

/** Whether an event renders as a banner span (all-day or crossing midnight)
 *  rather than a timed block inside one day column. */
export function isSpanEvent(ev: PCEvent): boolean {
	if (ev.allDay) return true;
	const s = eventDaySpan(ev);
	return s.startKey !== s.endKey;
}

/** The slice of a timed event inside one local day, in minutes since that
 *  day's midnight; null when the event does not touch the day. Minutes come
 *  from wall-clock components, so a DST-shifted day still draws sanely. */
export function clipToDay(ev: { startMs: number; endMs: number }, key: string): { startMin: number; endMin: number } | null {
	const dayStartMs = msOfKey(key);
	const dayEndMs = msOfKey(addDays(key, 1));
	if (ev.endMs <= dayStartMs || ev.startMs >= dayEndMs) return null;
	const startMin = ev.startMs <= dayStartMs ? 0 : minutesOfMs(ev.startMs);
	let endMin = ev.endMs >= dayEndMs ? 1440 : minutesOfMs(ev.endMs);
	if (endMin <= startMin) endMin = Math.min(1440, startMin + 15); // zero-length slivers stay visible
	return { startMin, endMin };
}

export function eventsOnDay(events: PCEvent[], key: string): PCEvent[] {
	return events.filter((ev) => {
		const s = eventDaySpan(ev);
		return dayDiff(s.startKey, key) >= 0 && dayDiff(key, s.endKey) >= 0;
	});
}

/** All-day banners first, then by start, longer first, then title. */
export function sortEvents(a: PCEvent, b: PCEvent): number {
	if (a.allDay !== b.allDay) return a.allDay ? -1 : 1;
	if (a.startMs !== b.startMs) return a.startMs - b.startMs;
	if (a.endMs !== b.endMs) return b.endMs - a.endMs;
	return a.title.localeCompare(b.title);
}

/* ---------- overlap layout ---------- */

export interface Packed {
	col: number;
	cols: number;
}

/** Side-by-side columns for overlapping timed blocks (the week/day grid).
 *  Blocks that overlap share a cluster; every block in a cluster is told the
 *  cluster's column count so widths line up. Zero-length input is the
 *  caller's problem: give slivers a display length first. */
export function packColumns(items: { startMin: number; endMin: number }[]): Packed[] {
	const order = items
		.map((_, i) => i)
		.sort((a, b) => items[a].startMin - items[b].startMin || items[b].endMin - items[a].endMin);
	const res: Packed[] = items.map(() => ({ col: 0, cols: 1 }));
	let cluster: number[] = [];
	let clusterEnd = -Infinity;
	const colEnd: number[] = [];
	const flush = () => {
		const n = colEnd.length || 1;
		for (const i of cluster) res[i].cols = n;
		cluster = [];
		colEnd.length = 0;
	};
	for (const i of order) {
		const it = items[i];
		if (cluster.length && it.startMin >= clusterEnd) flush();
		let col = colEnd.findIndex((e) => e <= it.startMin);
		if (col === -1) {
			col = colEnd.length;
			colEnd.push(0);
		}
		colEnd[col] = it.endMin;
		res[i].col = col;
		cluster.push(i);
		if (it.endMin > clusterEnd) clusterEnd = it.endMin;
	}
	flush();
	return res;
}

/** Greedy lane assignment for banner spans in a week row: first lane whose
 *  last span ends before this one starts. Indices are day columns, inclusive. */
export function packLanes(spans: { start: number; end: number }[]): number[] {
	const order = spans
		.map((_, i) => i)
		.sort((a, b) => spans[a].start - spans[b].start || spans[b].end - spans[a].end);
	const laneEnd: number[] = [];
	const lanes: number[] = spans.map(() => 0);
	for (const i of order) {
		const s = spans[i];
		let lane = laneEnd.findIndex((e) => e < s.start);
		if (lane === -1) {
			lane = laneEnd.length;
			laneEnd.push(-1);
		}
		laneEnd[lane] = s.end;
		lanes[i] = lane;
	}
	return lanes;
}

export interface RowSpan {
	ev: PCEvent;
	/** Day columns within the row, inclusive, clipped to 0..rowLen-1. */
	startIdx: number;
	endIdx: number;
	/** False when the event continues from before / past this row. */
	startsHere: boolean;
	endsHere: boolean;
	lane: number;
}

/** The banner spans of one week row, clipped, packed into lanes. */
export function spansForRow(events: PCEvent[], rowKeys: string[]): RowSpan[] {
	const first = rowKeys[0];
	const last = rowKeys[rowKeys.length - 1];
	const raw: Omit<RowSpan, "lane">[] = [];
	for (const ev of events) {
		if (!isSpanEvent(ev)) continue;
		const s = eventDaySpan(ev);
		if (dayDiff(s.endKey, first) > 0 || dayDiff(last, s.startKey) > 0) continue;
		raw.push({
			ev,
			startIdx: Math.max(0, dayDiff(first, s.startKey)),
			endIdx: Math.min(rowKeys.length - 1, dayDiff(first, s.endKey)),
			startsHere: dayDiff(first, s.startKey) >= 0,
			endsHere: dayDiff(first, s.endKey) <= rowKeys.length - 1,
		});
	}
	raw.sort((a, b) => a.startIdx - b.startIdx || b.endIdx - a.endIdx || sortEvents(a.ev, b.ev));
	const lanes = packLanes(raw.map((r) => ({ start: r.startIdx, end: r.endIdx })));
	return raw.map((r, i) => ({ ...r, lane: lanes[i] }));
}

/** Timed (non-banner) events starting on one day, for month cells. */
export function timedOnDay(events: PCEvent[], key: string): PCEvent[] {
	return events.filter((ev) => !isSpanEvent(ev) && keyOfMs(ev.startMs) === key).sort(sortEvents);
}

/** Agenda grouping: only days that have events, each day's list sorted. */
export function groupByDay(events: PCEvent[], fromKey: string, toKey: string): { key: string; events: PCEvent[] }[] {
	const out: { key: string; events: PCEvent[] }[] = [];
	for (let k = fromKey; dayDiff(k, toKey) >= 0; k = addDays(k, 1)) {
		const list = eventsOnDay(events, k).sort(sortEvents);
		if (list.length) out.push({ key: k, events: list });
	}
	return out;
}

/* ---------- view windows and navigation ---------- */

export type ViewMode = "month" | "week" | "workweek" | "day" | "agenda";

/** Monday through Friday of the week containing `key`. A weekend day belongs
 *  to the work week that just ended, matching how Outlook lands there. */
export function workWeekDays(key: string): string[] {
	return weekDays(key, true).slice(0, 5);
}

/** The inclusive day range a view needs events for. */
/** N consecutive day keys from the anchor: the day view's columns. */
export function daySpanKeys(anchorKey: string, count: number): string[] {
	return Array.from({ length: Math.max(1, Math.min(7, count)) }, (_, i) => addDays(anchorKey, i));
}

export function viewWindow(mode: ViewMode, anchorKey: string, weekStartsMonday: boolean, agendaDays = 30, dayCount = 1): { fromKey: string; toKey: string } {
	if (mode === "month") {
		const cells = monthGrid(+anchorKey.slice(0, 4), +anchorKey.slice(5, 7) - 1, weekStartsMonday);
		return { fromKey: cells[0].key, toKey: cells[41].key };
	}
	if (mode === "week") {
		const days = weekDays(anchorKey, weekStartsMonday);
		return { fromKey: days[0], toKey: days[6] };
	}
	if (mode === "workweek") {
		const days = workWeekDays(anchorKey);
		return { fromKey: days[0], toKey: days[4] };
	}
	if (mode === "day") {
		const days = daySpanKeys(anchorKey, dayCount);
		return { fromKey: days[0], toKey: days[days.length - 1] };
	}
	return { fromKey: anchorKey, toKey: addDays(anchorKey, Math.max(1, agendaDays) - 1) };
}

/** Where prev/next lands the anchor. Months move by calendar month and pin to
 *  the 1st so a Jan 31 anchor cannot skip February. */
export function stepAnchor(mode: ViewMode, anchorKey: string, dir: 1 | -1, agendaDays = 30, dayCount = 1): string {
	if (mode === "month") {
		let y = +anchorKey.slice(0, 4);
		let m0 = +anchorKey.slice(5, 7) - 1 + dir;
		if (m0 < 0) {
			m0 = 11;
			y--;
		}
		if (m0 > 11) {
			m0 = 0;
			y++;
		}
		return `${y}-${String(m0 + 1).padStart(2, "0")}-01`;
	}
	if (mode === "week" || mode === "workweek") return addDays(anchorKey, 7 * dir);
	if (mode === "day") return addDays(anchorKey, Math.max(1, Math.min(7, dayCount)) * dir);
	return addDays(anchorKey, Math.max(1, agendaDays) * dir);
}

/* ---------- labels ---------- */

export const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
export const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
export const DAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/** "9 AM", "9:30 AM", "12:05 PM" / 24h "09:00". `short` drops :00 in 12h. */
export function fmtClock(minutes: number, use24h: boolean, short = false): string {
	const h = Math.floor(minutes / 60) % 24;
	const m = minutes % 60;
	if (use24h) return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
	const ap = h < 12 ? "AM" : "PM";
	const h12 = h % 12 === 0 ? 12 : h % 12;
	if (short && m === 0) return `${h12} ${ap}`;
	return `${h12}:${String(m).padStart(2, "0")} ${ap}`;
}

export function fmtTimeOfMs(ms: number, use24h: boolean, short = false): string {
	return fmtClock(minutesOfMs(ms), use24h, short);
}

/** "Jul 17" / with year "Jul 17, 2026". */
export function fmtDayShort(key: string, withYear = false): string {
	const mo = MONTHS_SHORT[+key.slice(5, 7) - 1];
	const d = +key.slice(8, 10);
	return withYear ? `${mo} ${d}, ${key.slice(0, 4)}` : `${mo} ${d}`;
}

/** "Thursday, Jul 17" for agenda day headings. */
export function fmtDayHeading(key: string): string {
	return `${DAYS[dayOfWeek(key)]}, ${fmtDayShort(key)}`;
}

/** The header title for each mode. Ranges use a plain hyphen. */
export function periodLabel(mode: ViewMode, anchorKey: string, weekStartsMonday: boolean, agendaDays = 30, dayCount = 1): string {
	if (mode === "month") return `${MONTHS[+anchorKey.slice(5, 7) - 1]} ${anchorKey.slice(0, 4)}`;
	if (mode === "day" && dayCount <= 1) return `${DAYS[dayOfWeek(anchorKey)]}, ${fmtDayShort(anchorKey, true)}`;
	const { fromKey, toKey } = viewWindow(mode, anchorKey, weekStartsMonday, agendaDays, dayCount);
	const sameYear = fromKey.slice(0, 4) === toKey.slice(0, 4);
	const sameMonth = sameYear && fromKey.slice(5, 7) === toKey.slice(5, 7);
	if (sameMonth) return `${fmtDayShort(fromKey)} - ${+toKey.slice(8, 10)}, ${toKey.slice(0, 4)}`;
	if (sameYear) return `${fmtDayShort(fromKey)} - ${fmtDayShort(toKey)}, ${toKey.slice(0, 4)}`;
	return `${fmtDayShort(fromKey, true)} - ${fmtDayShort(toKey, true)}`;
}

/** "9:00 AM - 9:30 AM", multi-day "Jul 17, 9:00 PM - Jul 18, 7:00 AM",
 *  all-day "All day" or "Jul 17 - Jul 19". */
export function fmtEventRange(ev: PCEvent, use24h: boolean): string {
	const s = eventDaySpan(ev);
	if (ev.allDay) return s.startKey === s.endKey ? "All day" : `${fmtDayShort(s.startKey)} - ${fmtDayShort(s.endKey)}`;
	if (s.startKey === s.endKey) return `${fmtTimeOfMs(ev.startMs, use24h)} - ${fmtTimeOfMs(ev.endMs, use24h)}`;
	return `${fmtDayShort(s.startKey)}, ${fmtTimeOfMs(ev.startMs, use24h)} - ${fmtDayShort(s.endKey)}, ${fmtTimeOfMs(ev.endMs, use24h)}`;
}

/** A clock reading in another IANA zone; null when the zone name is invalid,
 *  which is also how callers decide whether to draw the column at all. */
export function fmtZoneClock(ms: number, tz: string, use24h: boolean): string | null {
	try {
		const fmt = use24h
			? new Intl.DateTimeFormat("en-US", { timeZone: tz, hour: "2-digit", minute: "2-digit", hourCycle: "h23" })
			: new Intl.DateTimeFormat("en-US", { timeZone: tz, hour: "numeric", minute: "2-digit", hour12: true });
		const s = fmt.format(new Date(ms));
		return use24h ? s : s.replace(":00 ", " ");
	} catch {
		return null;
	}
}

/** The timed events whose reminder should fire right now: lead time reached,
 *  not yet started more than a minute ago, not declined, not already fired.
 *  The caller owns the fired set and clears it on day change. */
export function dueReminders(events: PCEvent[], nowMs: number, leadMin: number, fired: Set<string>): PCEvent[] {
	return events.filter((ev) => {
		if (ev.allDay || ev.declined || fired.has(ev.id)) return false;
		return nowMs >= ev.startMs - leadMin * 60000 && nowMs < ev.startMs + 60000;
	});
}

/* ---------- mail (the triage inbox) ---------- */

export interface PCMail {
	id: string;
	accountId: string;
	accountLabel: string;
	from: string;
	fromAddress: string;
	subject: string;
	preview: string;
	receivedMs: number;
	unread: boolean;
	webLink?: string;
	hasAttachments?: boolean;
	/** The folder holding the message, when the fetch carried it; the Unread
	 *  search folder uses it to scope and label results. */
	folderId?: string;
	/** Marked high importance by the sender. */
	priority?: boolean;
	/** Flagged for follow-up in the mailbox. */
	flagged?: boolean;
	/** Addressed to the account directly (on the To line, not just Cc). */
	toMe?: boolean;
}

export interface GraphMailLike {
	id?: string;
	subject?: string;
	from?: { emailAddress?: { name?: string; address?: string } };
	receivedDateTime?: string;
	bodyPreview?: string;
	isRead?: boolean;
	webLink?: string;
	hasAttachments?: boolean;
	parentFolderId?: string;
	importance?: string;
	flag?: { flagStatus?: string };
	toRecipients?: { emailAddress?: { address?: string } }[];
}

/** One Graph inbox message to a PCMail; null when unreadable. `ownAddress`
 *  is the mailbox's own email, which decides the To-me flag. */
export function graphMailToPC(m: GraphMailLike, accountId: string, accountLabel: string, ownAddress?: string): PCMail | null {
	if (!m.id) return null;
	const receivedMs = Date.parse(m.receivedDateTime ?? "");
	if (!Number.isFinite(receivedMs)) return null;
	const own = (ownAddress ?? "").trim().toLowerCase();
	return {
		id: m.id,
		accountId,
		accountLabel,
		from: m.from?.emailAddress?.name || m.from?.emailAddress?.address || "(unknown)",
		fromAddress: m.from?.emailAddress?.address || "",
		subject: m.subject?.trim() || "(no subject)",
		preview: (m.bodyPreview ?? "").replace(/\s+/g, " ").trim(),
		receivedMs,
		unread: m.isRead === false,
		webLink: m.webLink || undefined,
		hasAttachments: !!m.hasAttachments,
		folderId: m.parentFolderId || undefined,
		priority: m.importance === "high" || undefined,
		flagged: m.flag?.flagStatus === "flagged" || undefined,
		toMe: (own && (m.toRecipients ?? []).some((r) => (r.emailAddress?.address ?? "").toLowerCase() === own)) || undefined,
	};
}

export interface MailFolder {
	id: string;
	name: string;
	parentId: string | null;
	unread: number;
	total: number;
	hasChildren: boolean;
}

export interface GraphFolderLike {
	id?: string;
	displayName?: string;
	parentFolderId?: string;
	childFolderCount?: number;
	unreadItemCount?: number;
	totalItemCount?: number;
}

export function graphFolderToPC(f: GraphFolderLike): MailFolder | null {
	if (!f.id) return null;
	return {
		id: f.id,
		name: f.displayName?.trim() || "(folder)",
		parentId: f.parentFolderId || null,
		unread: typeof f.unreadItemCount === "number" ? f.unreadItemCount : 0,
		total: typeof f.totalItemCount === "number" ? f.totalItemCount : 0,
		hasChildren: (f.childFolderCount ?? 0) > 0,
	};
}

/** The folder tree flattened for display: the inbox and its subtree first,
 *  then the rest of the top level, children alphabetical and indented by
 *  depth. A collapsed folder keeps its row and swallows its subtree.
 *  `expandable` reflects FETCHED children, so a chevron never opens onto
 *  nothing. A folder whose parent was not fetched surfaces at the top level
 *  rather than vanishing. */
export function orderFolderTree(folders: MailFolder[], inboxId: string | null, collapsed?: Set<string>): { folder: MailFolder; depth: number; expandable: boolean }[] {
	const ids = new Set(folders.map((f) => f.id));
	const byParent = new Map<string | null, MailFolder[]>();
	for (const f of folders) {
		const p = f.parentId && ids.has(f.parentId) ? f.parentId : null;
		const arr = byParent.get(p);
		if (arr) arr.push(f);
		else byParent.set(p, [f]);
	}
	const alpha = (a: MailFolder, b: MailFolder) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
	const top = (byParent.get(null) ?? []).slice().sort((a, b) => {
		const ai = a.id === inboxId ? 0 : 1;
		const bi = b.id === inboxId ? 0 : 1;
		return ai - bi || alpha(a, b);
	});
	const out: { folder: MailFolder; depth: number; expandable: boolean }[] = [];
	const walk = (f: MailFolder, depth: number) => {
		const kids = (byParent.get(f.id) ?? []).slice().sort(alpha);
		out.push({ folder: f, depth, expandable: kids.length > 0 });
		if (collapsed?.has(f.id)) return;
		for (const child of kids) walk(child, depth + 1);
	};
	for (const f of top) walk(f, 0);
	return out;
}

/** The ids of a folder and everything under it: how the Unread search folder
 *  scopes itself to the inbox and all its subfolders. Empty when the root is
 *  unknown. */
export function folderSubtreeIds(folders: MailFolder[], rootId: string | null): Set<string> {
	const out = new Set<string>();
	if (!rootId) return out;
	out.add(rootId);
	let grew = true;
	while (grew) {
		grew = false;
		for (const f of folders) {
			if (!out.has(f.id) && f.parentId && out.has(f.parentId)) {
				out.add(f.id);
				grew = true;
			}
		}
	}
	return out;
}

/** The span a calendar fetch should actually cover: the visible window plus
 *  generous padding either side, so stepping week to week or month to month
 *  lands on events already fetched instead of on the network. */
export function padFetchWindow(fromMs: number, toMs: number, backDays = 45, fwdDays = 90): { fromMs: number; toMs: number } {
	return { fromMs: fromMs - backDays * 86400000, toMs: toMs + fwdDays * 86400000 };
}

/** Fold a delta round into a cached list: changed messages replace their old
 *  selves or insert, removed ids drop, newest first, capped. This is what
 *  keeps a folder list current without ever re-downloading it. */
export function mergeDeltaMessages(existing: PCMail[], changed: PCMail[], removedIds: string[], cap = 50): PCMail[] {
	const removed = new Set(removedIds);
	const byId = new Map<string, PCMail>();
	for (const m of existing) if (!removed.has(m.id)) byId.set(m.id, m);
	for (const m of changed) if (!removed.has(m.id)) byId.set(m.id, m);
	return [...byId.values()].sort((a, b) => b.receivedMs - a.receivedMs).slice(0, cap);
}

/** Inbox-style timestamps: a clock today, a date this year, a full date past
 *  that. */
export function fmtMailTime(ms: number, todayKey: string, use24h: boolean): string {
	const key = keyOfMs(ms);
	if (key === todayKey) return fmtTimeOfMs(ms, use24h, true);
	return fmtDayShort(key, key.slice(0, 4) !== todayKey.slice(0, 4));
}

/** A mail subject as an event title: reply/forward prefixes gone. */
export function subjectToEventTitle(subject: string): string {
	return subject.replace(/^\s*((re|fw|fwd|aw|sv)\s*:\s*)+/i, "").trim() || subject.trim();
}

/* ---------- attendee availability (Graph getSchedule) ---------- */

export interface BusyRun {
	startMs: number;
	endMs: number;
	kind: "busy" | "tentative" | "oof";
}

/** Graph's availabilityView digit string ("0022...") as merged busy runs.
 *  0/4 (free, working elsewhere) do not block; 1 tentative, 2 busy, 3 oof. */
export function parseAvailabilityView(view: string, windowStartMs: number, intervalMin: number): BusyRun[] {
	const out: BusyRun[] = [];
	const kindOf = (ch: string): BusyRun["kind"] | null => (ch === "1" ? "tentative" : ch === "2" ? "busy" : ch === "3" ? "oof" : null);
	let cur: BusyRun | null = null;
	for (let i = 0; i < view.length; i++) {
		const k = kindOf(view[i]);
		const s = windowStartMs + i * intervalMin * 60000;
		if (k && cur && cur.kind === k && cur.endMs === s) cur.endMs = s + intervalMin * 60000;
		else if (k) {
			cur = { startMs: s, endMs: s + intervalMin * 60000, kind: k };
			out.push(cur);
		} else cur = null;
	}
	return out;
}

/** The worst overlap a draft slot has with someone's runs; null means free. */
export function slotConflict(runs: BusyRun[], startMs: number, endMs: number): BusyRun["kind"] | null {
	const rank = { tentative: 1, busy: 2, oof: 3 };
	let worst: BusyRun["kind"] | null = null;
	for (const r of runs) {
		if (r.endMs <= startMs || r.startMs >= endMs) continue;
		if (!worst || rank[r.kind] > rank[worst]) worst = r.kind;
	}
	return worst;
}

/* ---------- availability ---------- */

/** Open gaps inside one day's working window. Busy = timed events not
 *  declined and not marked free; all-day banners never block a slot. */
export function freeSlotsForDay(events: PCEvent[], key: string, fromMin: number, toMin: number, minGapMin = 30): { startMin: number; endMin: number }[] {
	const busy = eventsOnDay(events, key)
		.filter((ev) => !ev.allDay && !ev.declined && !ev.transparent)
		.map((ev) => clipToDay(ev, key))
		.filter((c): c is { startMin: number; endMin: number } => c != null)
		.sort((a, b) => a.startMin - b.startMin);
	const out: { startMin: number; endMin: number }[] = [];
	let cursor = fromMin;
	for (const b of busy) {
		if (b.startMin - cursor >= minGapMin) out.push({ startMin: cursor, endMin: Math.min(b.startMin, toMin) });
		cursor = Math.max(cursor, b.endMin);
		if (cursor >= toMin) break;
	}
	if (toMin - cursor >= minGapMin) out.push({ startMin: cursor, endMin: toMin });
	return out;
}

/** Paste-able availability text: one line per day, "free" for a fully open
 *  day, "booked" for none, otherwise the slots. */
export function fmtFreeSlots(days: { key: string; slots: { startMin: number; endMin: number }[] }[], fromMin: number, toMin: number, use24h: boolean): string {
	return days
		.map((d) => {
			const label = `${DAYS_SHORT[dayOfWeek(d.key)]} ${fmtDayShort(d.key)}`;
			if (!d.slots.length) return `${label}: booked`;
			if (d.slots.length === 1 && d.slots[0].startMin <= fromMin && d.slots[0].endMin >= toMin) return `${label}: free`;
			return `${label}: ${d.slots.map((s) => `${fmtClock(s.startMin, use24h, true)} - ${fmtClock(s.endMin, use24h, true)}`).join(", ")}`;
		})
		.join("\n");
}

/* ---------- note names ---------- */

/** A string safe as an Obsidian filename: no path or link syntax, one line. */
export function sanitizeName(s: string): string {
	return (
		s
			.replace(/[\\/:*?"<>|#^[\]]/g, " ")
			.replace(/\s+/g, " ")
			.trim()
			.slice(0, 120)
			.trim() || "Untitled"
	);
}

/** The note filename for an event: {{date}}, {{time}}, {{title}},
 *  {{calendar}} tokens, then sanitized as a whole. */
export function renderNoteName(template: string, ev: PCEvent, use24h: boolean): string {
	const s = eventDaySpan(ev);
	const name = (template || "{{date}} {{title}}")
		.replace(/\{\{date\}\}/g, s.startKey)
		.replace(/\{\{time\}\}/g, ev.allDay ? "" : fmtClock(minutesOfMs(ev.startMs), true).replace(":", "."))
		.replace(/\{\{title\}\}/g, ev.title)
		.replace(/\{\{calendar\}\}/g, ev.calendarName ?? "");
	return sanitizeName(name);
}

/** The `power-calendar` code block's little config language: "date: today"
 *  or an ISO day (absent means derive from the note's filename, else today),
 *  and "days: N". Anything unrecognized is ignored, never an error. */
export function parseAgendaBlock(source: string): { date: string | null; days: number } {
	let date: string | null = null;
	let days = 1;
	for (const line of source.split("\n")) {
		const m = line.match(/^\s*([a-z]+)\s*:\s*(.+?)\s*$/i);
		if (!m) continue;
		const key = m[1].toLowerCase();
		if (key === "date" && m[2].toLowerCase() !== "today") date = dateKeyOf(m[2]);
		if (key === "days") {
			const n = parseInt(m[2], 10);
			if (Number.isFinite(n)) days = Math.max(1, Math.min(31, n));
		}
	}
	return { date, days };
}

/* ---------- vault-notes source ---------- */

/** A frontmatter date (and optional end) as an event span. "2026-07-17" is an
 *  all-day; "2026-07-17T09:30" is timed (an hour when no end says otherwise);
 *  an end property stretches either kind. Null when there is no date. */
export function vaultDateSpan(raw: unknown, endRaw?: unknown): { startMs: number; endMs: number; allDay: boolean } | null {
	const s = String(raw ?? "");
	const key = dateKeyOf(s);
	if (!key) return null;
	const tm = s.match(/T(\d{1,2}):(\d{2})/);
	const startMin = tm ? Math.min(1439, +tm[1] * 60 + +tm[2]) : null;
	const es = endRaw == null ? "" : String(endRaw);
	const eKey = dateKeyOf(es);
	if (startMin == null) {
		const endKey = eKey && dayDiff(key, eKey) > 0 ? eKey : key;
		return { startMs: msOfKey(key), endMs: msOfKey(addDays(endKey, 1)), allDay: true };
	}
	const startMs = msOfKey(key) + startMin * 60000;
	const etm = es.match(/T(\d{1,2}):(\d{2})/);
	if (eKey && etm) {
		const endMs = msOfKey(eKey) + Math.min(1439, +etm[1] * 60 + +etm[2]) * 60000;
		if (endMs > startMs) return { startMs, endMs, allDay: false };
	}
	return { startMs, endMs: startMs + 3600000, allDay: false };
}

/** An instant back into frontmatter shape, keeping the note's date-only or
 *  datetime character. */
export function fmtVaultDate(ms: number, withTime: boolean): string {
	const key = keyOfMs(ms);
	if (!withTime) return key;
	const min = minutesOfMs(ms);
	return `${key}T${String(Math.floor(min / 60)).padStart(2, "0")}:${String(min % 60).padStart(2, "0")}`;
}

/* ---------- text utilities ---------- */

/** HTML to readable plain text: real line breaks for block ends, tags gone,
 *  the common entities decoded. For detail cards, not for fidelity. */
export function stripHtml(html: string): string {
	return html
		.replace(/<style[\s\S]*?<\/style>/gi, " ")
		.replace(/<script[\s\S]*?<\/script>/gi, " ")
		.replace(/<br\s*\/?>/gi, "\n")
		.replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, "\n")
		.replace(/<[^>]+>/g, " ")
		.replace(/&nbsp;/gi, " ")
		.replace(/&amp;/gi, "&")
		.replace(/&lt;/gi, "<")
		.replace(/&gt;/gi, ">")
		.replace(/&quot;/gi, '"')
		.replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
		.replace(/[ \t]+/g, " ")
		.replace(/ ?\n ?/g, "\n")
		.replace(/\n{3,}/g, "\n\n")
		.trim();
}

const JOIN_RE = /https:\/\/(?:teams\.microsoft\.com|teams\.live\.com|meet\.google\.com|[\w.-]*zoom\.us|[\w.-]*webex\.com|meet\.jit\.si)\/[^\s"'<>\])]+/i;

/** What to call the thing a join link opens. A reminder that prints the link
 *  itself is unreadable, a Zoom URL carries a meeting id and a password token
 *  and wraps over three lines, and nobody reads it anyway: the Join button is
 *  what gets clicked. The name is all the line has to say. An unrecognized host
 *  still beats the full URL, so it falls back to the bare domain. */
export function meetingProvider(url: string | undefined | null): string {
	if (!url) return "";
	const host = url.match(/^https?:\/\/([^/?#]+)/i)?.[1]?.toLowerCase();
	if (!host) return "";
	if (/(^|\.)zoom\.us$/.test(host)) return "Zoom";
	if (/(^|\.)teams\.(microsoft|live)\.com$/.test(host)) return "Teams";
	if (host === "meet.google.com") return "Google Meet";
	if (/(^|\.)webex\.com$/.test(host)) return "Webex";
	if (host === "meet.jit.si") return "Jitsi";
	return host.replace(/^www\./, "");
}

/** The first meeting-join link found in any of the given texts. */
export function findJoinUrl(...texts: (string | undefined | null)[]): string | null {
	for (const t of texts) {
		const m = t?.match(JOIN_RE);
		if (m) return m[0];
	}
	return null;
}

/** UTF-8 safe base64 (Buffer does not exist on Obsidian mobile). */
export function b64utf8(s: string): string {
	const bytes = new TextEncoder().encode(s);
	let bin = "";
	for (const b of bytes) bin += String.fromCharCode(b);
	return btoa(bin);
}

/** Outlook's search-folder catalog compiled to mailbox search queries (KQL).
 *  `param` carries the people, group, or category the type asks for, and
 *  `ownAddress` the mailbox's own email for the to-me type. */
export function searchFolderQuery(type: string, param: string, ownAddress: string): string {
	const people = param
		.split(/[,;]+/)
		.map((s) => s.trim())
		.filter(Boolean);
	switch (type) {
		case "unread":
			return "isread:false";
		case "flagged":
			return "isflagged:true";
		case "unreadOrFlagged":
			return "isread:false OR isflagged:true";
		case "important":
			return "importance:high";
		case "toMe":
			return `to:${ownAddress}`;
		case "fromPeople":
			return people.map((p) => `from:${p}`).join(" OR ");
		case "fromToPeople":
			return people.map((p) => `from:${p} OR to:${p}`).join(" OR ");
		case "toGroup":
			return people.map((p) => `to:${p}`).join(" OR ");
		case "category":
			return `category:"${param.trim()}"`;
		case "attachments":
			return "hasattachment:true";
		default:
			return param.trim();
	}
}

/* ---------- mail presentation ---------- */

/** Outlook-style avatar letters: the first letters of the first two words,
 *  or of an address's local-part segments ("uber-receipts@x" reads UR). */
export function avatarInitials(name: string): string {
	const n = name.trim();
	if (!n) return "?";
	const base = n.includes("@") ? n.split("@")[0] : n;
	const parts = base.split(/[\s._-]+/).filter(Boolean);
	if (!parts.length) return n[0].toUpperCase();
	return ((parts[0][0] ?? "") + (parts.length > 1 ? parts[1][0] ?? "" : "")).toUpperCase() || "?";
}

/** A stable palette color for a sender, hashed from the name. */
export function avatarColor(name: string): string {
	let h = 0;
	for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0;
	return paletteColor(Math.abs(h));
}

/* ---------- place lookup ---------- */

export interface GeoHit {
	name: string;
	admin1?: string;
	country_code?: string;
	latitude: number;
	longitude: number;
}

// prettier-ignore
const US_STATES: Record<string, string> = { AL: "Alabama", AK: "Alaska", AZ: "Arizona", AR: "Arkansas", CA: "California", CO: "Colorado", CT: "Connecticut", DE: "Delaware", DC: "District of Columbia", FL: "Florida", GA: "Georgia", HI: "Hawaii", ID: "Idaho", IL: "Illinois", IN: "Indiana", IA: "Iowa", KS: "Kansas", KY: "Kentucky", LA: "Louisiana", ME: "Maine", MD: "Maryland", MA: "Massachusetts", MI: "Michigan", MN: "Minnesota", MS: "Mississippi", MO: "Missouri", MT: "Montana", NE: "Nebraska", NV: "Nevada", NH: "New Hampshire", NJ: "New Jersey", NM: "New Mexico", NY: "New York", NC: "North Carolina", ND: "North Dakota", OH: "Ohio", OK: "Oklahoma", OR: "Oregon", PA: "Pennsylvania", RI: "Rhode Island", SC: "South Carolina", SD: "South Dakota", TN: "Tennessee", TX: "Texas", UT: "Utah", VT: "Vermont", VA: "Virginia", WA: "Washington", WV: "West Virginia", WI: "Wisconsin", WY: "Wyoming" };

/** "Flower Mound, TX" splits into the searchable city and the region hint;
 *  geocoders match on the bare name and reject the comma form whole. */
export function splitPlaceQuery(q: string): { name: string; region: string } {
	const i = q.indexOf(",");
	return i < 0 ? { name: q.trim(), region: "" } : { name: q.slice(0, i).trim(), region: q.slice(i + 1).trim() };
}

/** The candidate matching the region hint (a US state abbreviation, a state
 *  name, or a country code); the first hit when no hint or none match. */
export function pickGeoHit(region: string, hits: GeoHit[]): GeoHit | null {
	if (!hits.length) return null;
	const r = region.trim().toLowerCase();
	if (!r) return hits[0];
	const full = US_STATES[region.trim().toUpperCase()]?.toLowerCase();
	return (
		hits.find((h) => {
			const a1 = (h.admin1 ?? "").toLowerCase();
			return a1 === r || (!!full && a1 === full) || (h.country_code ?? "").toLowerCase() === r;
		}) ?? hits[0]
	);
}

/** WMO weather code to a compact glyph for the agenda's day headers. */
export function weatherGlyph(code: number): string {
	if (code === 0) return "☀️";
	if (code <= 2) return "🌤️";
	if (code === 3) return "☁️";
	if (code === 45 || code === 48) return "🌫️";
	if (code >= 51 && code <= 67) return "🌧️";
	if (code >= 71 && code <= 77) return "🌨️";
	if (code >= 80 && code <= 82) return "🌦️";
	if (code >= 85 && code <= 86) return "🌨️";
	if (code >= 95) return "⛈️";
	return "🌡️";
}

/* ---------- source colors ---------- */

export const PALETTE = ["#4c7bd9", "#d9564c", "#3fa564", "#c9832d", "#8a5cd9", "#2ea3b8", "#c94c8f", "#6b9e2e", "#b0623a", "#5c6bd9"];

export function paletteColor(i: number): string {
	return PALETTE[((i % PALETTE.length) + PALETTE.length) % PALETTE.length];
}

/* ---------- write-back geometry ---------- */

export function snapMin(min: number, step = 15): number {
	return Math.round(min / step) * step;
}

/** Where a dragged block lands: `move` carries the duration to the target day
 *  at the snapped new start; `resize` keeps the start and drags the end, never
 *  under 15 minutes. Deltas are raw pointer minutes; snapping happens here so
 *  every caller gets the same feel. */
export function dragTimes(
	startMs: number,
	endMs: number,
	targetDayKey: string,
	deltaMin: number,
	mode: "move" | "resize"
): { startMs: number; endMs: number } {
	if (mode === "resize") {
		const dur = Math.max(15, snapMin((endMs - startMs) / 60000 + deltaMin));
		return { startMs, endMs: startMs + dur * 60000 };
	}
	const startMin = Math.max(0, Math.min(1439, snapMin(minutesOfMs(startMs) + deltaMin)));
	const s = msOfKey(targetDayKey) + startMin * 60000;
	return { startMs: s, endMs: s + (endMs - startMs) };
}

/** Local wall clock of an instant, Graph-shaped: "2026-07-17T09:00:00". */
export function wallOfMs(ms: number): string {
	const d = new Date(ms);
	const p = (n: number) => String(n).padStart(2, "0");
	return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

export type RepeatKind = "none" | "daily" | "weekdays" | "weekly" | "monthly" | "yearly";

export interface EventDraft {
	title: string;
	startMs: number;
	/** Exclusive, like PCEvent. All-day drafts run midnight to midnight. */
	endMs: number;
	allDay: boolean;
	location?: string;
	description?: string;
	/** People to invite; null/undefined leaves an event's attendees untouched. */
	attendees?: { name?: string; email: string }[] | null;
	/** Recurrence at creation time only; edits never touch a series' rule. */
	repeat?: RepeatKind;
	/** How the slot reads to others; undefined leaves it alone (busy on create). */
	showAs?: "busy" | "free" | "tentative";
}

/** "a@x.com, Bob <b@x.com>; c@x.com" into invitees. Entries without an @ are
 *  dropped rather than guessed at. */
export function parseAttendeeInput(s: string): { name?: string; email: string }[] {
	const out: { name?: string; email: string }[] = [];
	for (const part of s.split(/[,;\n]+/)) {
		const t = part.trim();
		if (!t) continue;
		const m = t.match(/^(.*?)\s*<([^>]+)>$/);
		const email = (m ? m[2] : t).trim();
		if (!email.includes("@")) continue;
		const name = m?.[1].trim();
		out.push(name ? { name, email } : { email });
	}
	return out;
}

const GRAPH_DOW = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
const RRULE_DOW = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"];

/** Graph's recurrence object for a simple repeat rule anchored on startKey. */
export function graphRecurrence(kind: RepeatKind, startKey: string): Record<string, unknown> | null {
	if (kind === "none") return null;
	const range = { type: "noEnd", startDate: startKey };
	if (kind === "daily") return { pattern: { type: "daily", interval: 1 }, range };
	if (kind === "weekdays") return { pattern: { type: "weekly", interval: 1, daysOfWeek: GRAPH_DOW.slice(1, 6) }, range };
	if (kind === "weekly") return { pattern: { type: "weekly", interval: 1, daysOfWeek: [GRAPH_DOW[dayOfWeek(startKey)]] }, range };
	if (kind === "monthly") return { pattern: { type: "absoluteMonthly", interval: 1, dayOfMonth: +startKey.slice(8, 10) }, range };
	return { pattern: { type: "absoluteYearly", interval: 1, dayOfMonth: +startKey.slice(8, 10), month: +startKey.slice(5, 7) }, range };
}

/** Google's RRULE line(s) for the same simple rules. */
export function googleRecurrence(kind: RepeatKind, startKey: string): string[] | null {
	if (kind === "none") return null;
	if (kind === "daily") return ["RRULE:FREQ=DAILY"];
	if (kind === "weekdays") return ["RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR"];
	if (kind === "weekly") return [`RRULE:FREQ=WEEKLY;BYDAY=${RRULE_DOW[dayOfWeek(startKey)]}`];
	if (kind === "monthly") return ["RRULE:FREQ=MONTHLY"];
	return ["RRULE:FREQ=YEARLY"];
}

/** The start/end pair every Graph write shares. All-day events must land on
 *  midnights of the given zone, which local wall clocks of local midnights are. */
export function graphTimesBody(startMs: number, endMs: number, allDay: boolean, tz: string): Record<string, unknown> {
	return {
		isAllDay: allDay,
		start: { dateTime: wallOfMs(startMs), timeZone: tz },
		end: { dateTime: wallOfMs(Math.max(endMs, startMs + (allDay ? 86400000 : 60000))), timeZone: tz },
	};
}

/** A full create/update body for Graph's events endpoints. */
export function graphEventBody(d: EventDraft, tz: string): Record<string, unknown> {
	const body: Record<string, unknown> = {
		subject: d.title.trim() || "(no title)",
		...graphTimesBody(d.startMs, d.endMs, d.allDay, tz),
	};
	if (d.location != null) body.location = { displayName: d.location.trim() };
	if (d.description != null) body.body = { contentType: "text", content: d.description };
	if (d.attendees != null) body.attendees = d.attendees.map((a) => ({ emailAddress: { address: a.email, ...(a.name ? { name: a.name } : {}) }, type: "required" }));
	const rec = d.repeat ? graphRecurrence(d.repeat, keyOfMs(d.startMs)) : null;
	if (rec) body.recurrence = rec;
	if (d.showAs) body.showAs = d.showAs;
	return body;
}

/** Attendees without the organizer's duplicate entry (Graph lists organizers
 *  among their own attendees, which reads twice on a card). */
export function dedupePeople(organizer: string | undefined, attendees: string[] | undefined): string[] {
	const org = organizer?.trim().toLowerCase();
	return (attendees ?? []).filter((a) => a.trim().toLowerCase() !== org);
}

/* ---------- Microsoft Graph event mapping ---------- */

export interface GraphDateTime {
	dateTime?: string;
	timeZone?: string;
}

export interface GraphEventLike {
	id?: string;
	subject?: string;
	isAllDay?: boolean;
	isCancelled?: boolean;
	showAs?: string;
	type?: string;
	start?: GraphDateTime;
	end?: GraphDateTime;
	location?: { displayName?: string };
	organizer?: { emailAddress?: { name?: string; address?: string } };
	attendees?: { type?: string; emailAddress?: { name?: string; address?: string } }[];
	onlineMeeting?: { joinUrl?: string } | null;
	isOnlineMeeting?: boolean;
	webLink?: string;
	bodyPreview?: string;
	responseStatus?: { response?: string };
	isOrganizer?: boolean;
	seriesMasterId?: string;
	categories?: string[];
}

/** "2026-07-17T09:00:00.0000000" read as LOCAL wall time. calendarView is
 *  asked to answer in the device's zone, so the wall clock IS local time. */
export function parseWallClock(s: string): number | null {
	const m = s.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
	if (!m) return null;
	return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime();
}

/** One Graph calendarView item to a PCEvent; null for canceled or unreadable
 *  items. calendarView expands recurrences server-side, so every item here is
 *  already a concrete instance. */
export function graphToPC(ev: GraphEventLike, src: { sourceId: string; calendarName?: string; color?: string; writable?: boolean }): PCEvent | null {
	if (ev.isCancelled) return null;
	const startMs = parseWallClock(ev.start?.dateTime ?? "");
	if (startMs == null) return null;
	let endMs = parseWallClock(ev.end?.dateTime ?? "") ?? startMs;
	if (endMs < startMs) endMs = startMs;
	const people = (ev.attendees ?? []).filter((a) => a.type !== "resource"); // rooms and equipment are not people
	const attendees = people.map((a) => a.emailAddress?.name || a.emailAddress?.address || "").filter(Boolean);
	const attendeeDetail = people
		.map((a) => ({ ...(a.emailAddress?.name ? { name: a.emailAddress.name } : {}), ...(a.emailAddress?.address ? { email: a.emailAddress.address } : {}) }))
		.filter((a) => a.name || a.email);
	const resp = ev.responseStatus?.response;
	const myResponse: PCEvent["myResponse"] = resp === "accepted" ? "accepted" : resp === "tentativelyAccepted" ? "tentative" : resp === "declined" ? "declined" : "none";
	return {
		id: ev.id || `${startMs}:${ev.subject ?? ""}`,
		sourceId: src.sourceId,
		calendarName: src.calendarName,
		color: src.color,
		title: ev.subject?.trim() || "(no title)",
		allDay: !!ev.isAllDay,
		startMs,
		endMs,
		location: ev.location?.displayName || undefined,
		organizer: ev.organizer?.emailAddress?.name || ev.organizer?.emailAddress?.address || undefined,
		attendees: attendees.length ? attendees : undefined,
		description: ev.bodyPreview || undefined,
		url: ev.webLink || undefined,
		joinUrl: ev.onlineMeeting?.joinUrl || findJoinUrl(ev.bodyPreview, ev.location?.displayName) || undefined,
		recurring: ev.type === "occurrence" || ev.type === "exception",
		seriesId: ev.seriesMasterId || undefined,
		categories: ev.categories?.length ? ev.categories : undefined,
		tentative: ev.showAs === "tentative",
		declined: ev.responseStatus?.response === "declined",
		canEdit: !!src.writable && ev.isOrganizer !== false,
		attendeeDetail: attendeeDetail.length ? attendeeDetail : undefined,
		myResponse,
		canRsvp: !!src.writable && ev.isOrganizer === false && people.length > 0,
		transparent: ev.showAs === "free",
	};
}

/** A sign-in failure's actual fix, when the AADSTS code is a known setup
 *  problem (single-tenant app on 'common', wrong client ID, public client
 *  flows off, missing consent). Null when the message carries no known code. */
/** Invite bodies end in a wall of join boilerplate (Teams, Zoom, Meet); the
 *  event card shows the human part only. Cuts from the first separator line
 *  or platform block header to the end. */
export function stripMeetingBoilerplate(text: string): string {
	const lines = text.split(/\r?\n/);
	const cut = lines.findIndex((l) => /^\s*_{6,}\s*$/.test(l) || /^(Microsoft Teams meeting|Join Zoom Meeting|Join on your computer|Google Meet joining info)/i.test(l.trim()));
	return (cut >= 0 ? lines.slice(0, cut) : lines).join("\n").trim();
}

export function graphSetupHint(text: string): string | null {
	if (/AADSTS50059|AADSTS50194/.test(text))
		return "Your Azure app is registered for one organization only, so 'common' cannot be used: paste the Directory (tenant) ID from the app's Overview page into the Tenant field.";
	if (/AADSTS700016/.test(text))
		return "The Application (client) ID was not found in this tenant: check both the ID and the Tenant field against the app's Overview page.";
	if (/AADSTS7000218/.test(text)) return "Turn on 'Allow public client flows' under the app's Authentication settings.";
	if (/AADSTS500200|AADSTS50020/.test(text))
		return "This account cannot sign in through this app registration: either the app does not accept personal Microsoft accounts, or the account belongs to a different organization than the Tenant. A personal outlook.com account connects through Add account, picked as Personal, with its own app registration; the wizard can create the app on the spot.";
	if (/ErrorAccessDenied|Access is denied/i.test(text)) return "This account's sign-in predates a permission this feature needs: press Reconnect on the account in settings to grant it.";
	if (/AADSTS65001|consent/i.test(text) && /AADSTS/.test(text))
		return "The app is missing consent: grant the delegated Calendars.ReadWrite, Mail.ReadWrite, and Mail.Send permissions (or have an admin grant consent) under API permissions.";
	return null;
}

/** The claims of a JWT's payload segment, or null when it does not parse.
 *  Sign-ins return an id_token whose claims name the account (email,
 *  preferred_username), which is how account rows get their labels without
 *  asking for any extra permission. */
export function decodeJwtPayload(jwt: string): Record<string, unknown> | null {
	const part = jwt.split(".")[1];
	if (!part) return null;
	try {
		const b64 = part.replace(/-/g, "+").replace(/_/g, "/");
		const bin = atob(b64 + "=".repeat((4 - (b64.length % 4)) % 4));
		const bytes = Uint8Array.from(bin, (ch) => ch.charCodeAt(0));
		return JSON.parse(new TextDecoder().decode(bytes)) as Record<string, unknown>;
	} catch {
		return null;
	}
}

/* ---------- Google Calendar mapping ---------- */

export interface GoogleEventLike {
	id?: string;
	status?: string;
	summary?: string;
	start?: { date?: string; dateTime?: string; timeZone?: string };
	end?: { date?: string; dateTime?: string; timeZone?: string };
	location?: string;
	description?: string;
	organizer?: { email?: string; displayName?: string; self?: boolean };
	attendees?: { email?: string; displayName?: string; responseStatus?: string; self?: boolean; resource?: boolean }[];
	hangoutLink?: string;
	htmlLink?: string;
	recurringEventId?: string;
	transparency?: string;
}

/** One Google events item (singleEvents expansion) to a PCEvent; null for
 *  canceled or unreadable items. Timed values carry an RFC3339 offset, so
 *  Date.parse lands the exact instant; date-only values are local days. */
export function googleToPC(ev: GoogleEventLike, src: { sourceId: string; calendarName?: string; color?: string; writable?: boolean }): PCEvent | null {
	if (ev.status === "cancelled") return null;
	let startMs: number;
	let endMs: number;
	let allDay = false;
	if (ev.start?.date) {
		allDay = true;
		startMs = msOfKey(ev.start.date);
		endMs = ev.end?.date ? msOfKey(ev.end.date) : startMs + 86400000; // end.date is already exclusive
	} else {
		startMs = Date.parse(ev.start?.dateTime ?? "");
		endMs = ev.end?.dateTime ? Date.parse(ev.end.dateTime) : startMs;
	}
	if (!Number.isFinite(startMs)) return null;
	if (!Number.isFinite(endMs) || endMs < startMs) endMs = startMs;
	const self = ev.attendees?.find((a) => a.self);
	const people = (ev.attendees ?? []).filter((a) => !a.resource);
	const attendees = people.map((a) => a.displayName || a.email || "").filter(Boolean);
	const attendeeDetail = people
		.map((a) => ({ ...(a.displayName ? { name: a.displayName } : {}), ...(a.email ? { email: a.email } : {}) }))
		.filter((a) => a.name || a.email);
	const selfResp = self?.responseStatus;
	const myResponse: PCEvent["myResponse"] = selfResp === "accepted" ? "accepted" : selfResp === "tentative" ? "tentative" : selfResp === "declined" ? "declined" : "none";
	const rawDesc = ev.description?.trim();
	return {
		id: ev.id || `${startMs}:${ev.summary ?? ""}`,
		sourceId: src.sourceId,
		calendarName: src.calendarName,
		color: src.color,
		title: ev.summary?.trim() || "(no title)",
		allDay,
		startMs,
		endMs,
		location: ev.location || undefined,
		organizer: ev.organizer?.displayName || ev.organizer?.email || undefined,
		attendees: attendees.length ? attendees : undefined,
		description: rawDesc ? (/<\w+[^>]*>/.test(rawDesc) ? stripHtml(rawDesc) : rawDesc) : undefined,
		url: ev.htmlLink || undefined,
		joinUrl: ev.hangoutLink || findJoinUrl(ev.location, ev.description) || undefined,
		recurring: !!ev.recurringEventId,
		seriesId: ev.recurringEventId || undefined,
		tentative: ev.status === "tentative" || self?.responseStatus === "tentative",
		declined: self?.responseStatus === "declined",
		canEdit: !!src.writable,
		attendeeDetail: attendeeDetail.length ? attendeeDetail : undefined,
		myResponse,
		canRsvp: !!src.writable && !!self && !ev.organizer?.self,
		transparent: ev.transparency === "transparent",
	};
}

/** Google's start/end pair: all-day uses exclusive date fields, timed uses a
 *  local wall clock plus the IANA zone (RFC3339 offsets become unnecessary
 *  when the zone rides along). */
export function googleTimesBody(startMs: number, endMs: number, allDay: boolean, tz: string): Record<string, unknown> {
	if (allDay) return { start: { date: keyOfMs(startMs) }, end: { date: keyOfMs(Math.max(endMs, startMs + 86400000)) } };
	return {
		start: { dateTime: wallOfMs(startMs), timeZone: tz },
		end: { dateTime: wallOfMs(Math.max(endMs, startMs + 60000)), timeZone: tz },
	};
}

export function googleEventBody(d: EventDraft, tz: string): Record<string, unknown> {
	const body: Record<string, unknown> = {
		summary: d.title.trim() || "(no title)",
		...googleTimesBody(d.startMs, d.endMs, d.allDay, tz),
	};
	if (d.location != null) body.location = d.location.trim();
	if (d.description != null) body.description = d.description;
	if (d.attendees != null) body.attendees = d.attendees.map((a) => ({ email: a.email, ...(a.name ? { displayName: a.name } : {}) }));
	const rec = d.repeat ? googleRecurrence(d.repeat, keyOfMs(d.startMs)) : null;
	if (rec) body.recurrence = rec;
	// Google has no showAs: free maps to transparency, tentative to status
	if (d.showAs) {
		body.transparency = d.showAs === "free" ? "transparent" : "opaque";
		body.status = d.showAs === "tentative" ? "tentative" : "confirmed";
	}
	return body;
}

/* ---------- settings persistence ---------- */

/**
 * Merge our settings over what is on disk RIGHT NOW, for a save.
 *
 * data.json is a synced file: other devices write it, and a device that has been
 * idle still holds whatever it read when its plugin loaded. Writing that whole
 * object back reverts every change made anywhere else since. With sign-in tokens
 * living in this file, a stale overwrite does not just lose a preference, it
 * logs another device out.
 *
 * So a save may only carry the keys we actually changed. `baseline` is the state
 * we last read from or wrote to disk, so anything differing from it is ours:
 * those keys overwrite. Every untouched key takes the disk's value. A key absent
 * from disk was written by a version that did not know it, and keeps ours rather
 * than resetting to a default.
 */
export function mergeForSave<T extends object>(ours: T, baseline: T, disk: Partial<T> | null): T {
	const out = { ...ours };
	if (!disk) return out;
	for (const k of Object.keys(ours) as (keyof T)[]) {
		if (!(k in disk)) continue; // disk has never heard of this key; ours stands
		const o = ours[k];
		const b = baseline[k];
		const d = disk[k];
		if (isRecord(o) && isRecord(b) && isRecord(d)) {
			out[k] = mergeEntries(o, b, d) as T[keyof T];
			continue;
		}
		const changedByUs = JSON.stringify(o) !== JSON.stringify(b);
		if (!changedByUs) out[k] = d as T[keyof T];
	}
	return out;
}

/** A per-item map, as opposed to a value that means something whole. Arrays are
 *  values here: a list's order and membership are the thing itself. */
function isRecord(v: unknown): v is Record<string, unknown> {
	return typeof v === "object" && v !== null && !Array.isArray(v);
}

/**
 * The same three-way rule, entry by entry.
 *
 * A key holding one value per item (per folder, per field, per speaker) is a
 * whole vault's worth of settings behind a single name, and merging it whole
 * meant changing ONE of them published all of them. Every item another device
 * configured since this one last read was erased by a device that had never
 * seen it.
 *
 * Start from the disk, so anything another device set survives; drop only what
 * we deliberately removed (present in the baseline, gone from ours); then lay
 * our own changed entries over the top. Two devices editing the SAME item still
 * settles last-writer-wins, but that is one item losing a race rather than
 * everything losing it.
 */
function mergeEntries(
	ours: Record<string, unknown>,
	baseline: Record<string, unknown>,
	disk: Record<string, unknown>
): Record<string, unknown> {
	const out: Record<string, unknown> = {};
	for (const k of Object.keys(disk)) {
		const removedByUs = k in baseline && !(k in ours);
		if (!removedByUs) out[k] = disk[k];
	}
	for (const k of Object.keys(ours)) {
		const changedByUs = JSON.stringify(ours[k]) !== JSON.stringify(baseline[k]);
		if (changedByUs || !(k in disk)) out[k] = ours[k];
	}
	return out;
}

/* ---------------- Graph $batch ---------------- */

/** Graph refuses a $batch carrying more than twenty sub-requests. */
export const GRAPH_BATCH_MAX = 20;

/** Split a list into runs of at most `size`. */
export function chunk<T>(xs: readonly T[], size: number): T[][] {
	const n = Math.max(1, Math.floor(size));
	const out: T[][] = [];
	for (let i = 0; i < xs.length; i += n) out.push(xs.slice(i, i + n));
	return out;
}

export interface BatchRequest {
	id: string;
	method: string;
	url: string;
}

/** The body of a $batch that reads several messages in one round trip.
 *
 *  Sub-requests are keyed by their index rather than the message id: Graph ids
 *  are long, opaque, and not guaranteed safe as a batch key, and the index maps
 *  back cleanly as long as the caller keeps the same array. */
export function buildMessageBatch(ids: readonly string[], select: string): { requests: BatchRequest[] } {
	return {
		requests: ids.map((id, i) => ({
			id: String(i),
			method: "GET",
			url: `/me/messages/${encodeURIComponent(id)}?$select=${select}`,
		})),
	};
}

export interface BatchOutcome {
	ok: Map<string, Record<string, unknown>>;
	failed: string[];
	/** Longest Retry-After seen across throttled sub-responses, in ms. */
	retryAfterMs: number;
}

/** Sort a $batch reply back into successes and failures.
 *
 *  Three things make this less trivial than it looks: responses may come back
 *  in any order, a single sub-request can fail while its neighbours succeed,
 *  and Graph may omit a response entirely. Anything not accounted for is
 *  reported failed so the caller retries it rather than silently losing a
 *  message. */
export function parseMessageBatch(json: unknown, ids: readonly string[]): BatchOutcome {
	const ok = new Map<string, Record<string, unknown>>();
	const failed: string[] = [];
	let retryAfterMs = 0;
	const responses = (json as { responses?: unknown } | null)?.responses;
	if (!Array.isArray(responses)) return { ok, failed: [...ids], retryAfterMs };

	for (const raw of responses) {
		const r = raw as { id?: unknown; status?: unknown; body?: unknown; headers?: Record<string, string> };
		const i = parseInt(String(r.id ?? ""), 10);
		const id = Number.isInteger(i) ? ids[i] : undefined;
		if (!id) continue;
		const status = typeof r.status === "number" ? r.status : 0;
		if (status >= 200 && status < 400 && r.body && typeof r.body === "object") {
			ok.set(id, r.body as Record<string, unknown>);
			continue;
		}
		failed.push(id);
		if (status === 429 || status === 503) {
			const h = r.headers ?? {};
			const raw2 = h["Retry-After"] ?? h["retry-after"] ?? "";
			const secs = parseFloat(String(raw2));
			if (isFinite(secs) && secs > 0) retryAfterMs = Math.max(retryAfterMs, secs * 1000);
		}
	}
	const seen = new Set<string>([...ok.keys(), ...failed]);
	for (const id of ids) if (!seen.has(id)) failed.push(id);
	return { ok, failed, retryAfterMs };
}
