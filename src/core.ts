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
	/** Outlook categories on the message, by display name. */
	categories?: string[];
	/** Outlook's own Focused verdict for the message. Absent on anything
	 *  cached before the split inbox shipped, which counts as focused so
	 *  nothing quietly disappears into a lower section. */
	focused?: boolean;
	/** The Graph conversation this belongs to, when the fetch carried one.
	 *  Messages cached before threading shipped have none, which simply makes
	 *  each of them a conversation of one until the next refresh. */
	conversationId?: string;
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
	conversationId?: string;
	inferenceClassification?: string;
	categories?: string[];
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
		conversationId: m.conversationId || undefined,
		focused: m.inferenceClassification ? m.inferenceClassification === "focused" : undefined,
		categories: m.categories?.length ? m.categories : undefined,
	};
}

/* ---------- printing ---------- */

export interface PrintOptions {
	/** Body size in points; the whole document scales from it. */
	fontPt?: number;
	landscape?: boolean;
}

/** The @page rule for a printable document, so orientation and margins are
 *  decided in one place rather than in four stylesheets. */
export function pageRule(o: PrintOptions | undefined, defaultLandscape: boolean, marginMm: number): string {
	const landscape = o?.landscape ?? defaultLandscape;
	return `@page { ${landscape ? "size: landscape; " : ""}margin: ${marginMm}mm; }`;
}

/** Sizes as a scale rather than a point value, because the right size for a
 *  message is not the right size for a month grid: each style keeps its own
 *  base and the scale moves it. "Normal" is exactly what printed before any
 *  of this existed. */
export const PRINT_SCALES = [
	{ id: "xs", label: "Smallest", factor: 0.75 },
	{ id: "s", label: "Smaller", factor: 0.875 },
	{ id: "m", label: "Normal", factor: 1 },
	{ id: "l", label: "Larger", factor: 1.15 },
	{ id: "xl", label: "Largest", factor: 1.35 },
] as const;

export type PrintScaleId = (typeof PRINT_SCALES)[number]["id"];

/** A style's base size moved by one of the scale steps, to a tenth of a point
 *  so the stylesheet never carries a number nobody can read. */
export function scaledPt(basePt: number, scale: string): number {
	const f = PRINT_SCALES.find((s) => s.id === scale)?.factor ?? 1;
	return Math.round(basePt * f * 10) / 10;
}

const escHtml = (s: string) => (s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** A message as a standalone document to hand a printer.
 *
 *  Built as its own page rather than printing the window, because printing
 *  the window prints Obsidian: the sidebar, the folder tree, the message
 *  list, and a sliver of the mail. The header fields are escaped even though
 *  they came from a mailbox, since a subject is attacker-controlled text and
 *  this is the one place it stops being data and becomes markup. The body is
 *  already sanitized by the caller and goes in whole. */
export function printableHtml(m: { subject: string; from: string; to: string; date: string; bodyHtml: string; plain?: boolean }, o?: PrintOptions): string {
	const body = m.plain ? `<pre class="plain">${escHtml(m.bodyHtml)}</pre>` : m.bodyHtml;
	const row = (label: string, value: string) => (value ? `<tr><th>${escHtml(label)}</th><td>${escHtml(value)}</td></tr>` : "");
	return `<!doctype html><html><head><meta charset="utf-8"><title>${escHtml(m.subject)}</title><style>
:root { color-scheme: light; }
body { margin: 0; padding: 24px; background: #fff; color: #111;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: ${o?.fontPt ?? 12}pt; line-height: 1.45; }
h1 { font-size: 16pt; margin: 0 0 12px; }
table.head { border-collapse: collapse; margin: 0 0 16px; }
table.head th { text-align: left; padding: 1px 12px 1px 0; font-weight: 600; color: #555; vertical-align: top; white-space: nowrap; }
table.head td { padding: 1px 0; }
hr { border: 0; border-top: 1px solid #ccc; margin: 0 0 16px; }
img { max-width: 100%; height: auto; }
pre.plain { white-space: pre-wrap; word-wrap: break-word; font-family: inherit; font-size: inherit; margin: 0; }
table { max-width: 100%; }
a { color: #0645ad; }
${pageRule(o, false, 14)}
</style></head><body>
<h1>${escHtml(m.subject)}</h1>
<table class="head">${row("From", m.from)}${row("To", m.to)}${row("Date", m.date)}</table>
<hr>
${body}
</body></html>`;
}

export interface PrintDay {
	heading: string;
	events: { when: string; title: string; where?: string }[];
}

/** A stretch of calendar as an agenda, Outlook's Calendar Details style.
 *
 *  A list rather than a reproduction of the timed grid: a grid printed at
 *  page width turns an hour into four millimetres and a day of meetings into
 *  a smear, while the list says what is actually happening and reads at arm's
 *  length. Empty days are kept, because "nothing on Thursday" is information. */
export function printableAgendaHtml(title: string, days: readonly PrintDay[], o?: PrintOptions): string {
	const rows = days
		.map((d) => {
			const evs = d.events.length
				? d.events
						.map((e) => `<tr><td class="nw t">${escHtml(e.when)}</td><td>${escHtml(e.title)}${e.where ? `<div class="where">${escHtml(e.where)}</div>` : ""}</td></tr>`)
						.join("")
				: `<tr><td class="nw t"></td><td class="none">Nothing scheduled</td></tr>`;
			return `<tr><th colspan="2" class="day">${escHtml(d.heading)}</th></tr>${evs}`;
		})
		.join("");
	return `<!doctype html><html><head><meta charset="utf-8"><title>${escHtml(title)}</title><style>
:root { color-scheme: light; }
body { margin: 0; padding: 20px; background: #fff; color: #111;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: ${o?.fontPt ?? 10.5}pt; }
h1 { font-size: 15pt; margin: 0 0 14px; }
table { border-collapse: collapse; width: 100%; }
th.day { text-align: left; font-size: 10pt; padding: 12px 0 3px; border-bottom: 1.5px solid #999; }
td { padding: 3px 8px 3px 0; border-bottom: 1px solid #eee; vertical-align: top; }
td.t { width: 130px; color: #444; }
td.nw { white-space: nowrap; }
td.none { color: #888; font-style: italic; }
div.where { color: #666; font-size: 9pt; }
tr { page-break-inside: avoid; }
${pageRule(o, false, 13)}
</style></head><body>
<h1>${escHtml(title)}</h1>
<table>${rows}</table>
</body></html>`;
}

export interface PrintCell {
	label: string;
	dim: boolean;
	events: { when: string; title: string }[];
}

/** A month as the grid it is on screen. Landscape by default, because a
 *  seven-column month squeezed into portrait is unreadable and the printer
 *  should be told rather than left to guess. */
export function printableMonthHtml(title: string, weekdays: readonly string[], weeks: readonly (readonly PrintCell[])[], o?: PrintOptions): string {
	const head = weekdays.map((d) => `<th>${escHtml(d)}</th>`).join("");
	const body = weeks
		.map(
			(w) =>
				`<tr>${w
					.map(
						(c) =>
							`<td class="${c.dim ? "dim" : ""}"><div class="num">${escHtml(c.label)}</div>${c.events
								.map((e) => `<div class="ev"><span class="t">${escHtml(e.when)}</span> ${escHtml(e.title)}</div>`)
								.join("")}</td>`
					)
					.join("")}</tr>`
		)
		.join("");
	return `<!doctype html><html><head><meta charset="utf-8"><title>${escHtml(title)}</title><style>
:root { color-scheme: light; }
body { margin: 0; padding: 14px; background: #fff; color: #111;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: ${o?.fontPt ?? 8.5}pt; }
h1 { font-size: 14pt; margin: 0 0 8px; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th { border: 1px solid #999; padding: 3px; font-size: 8pt; text-transform: uppercase; letter-spacing: 0.03em; background: #f2f2f2; }
td { border: 1px solid #bbb; padding: 3px; height: 92px; vertical-align: top; overflow: hidden; }
td.dim { background: #fafafa; color: #999; }
div.num { font-weight: 700; font-size: 9pt; margin-bottom: 2px; }
div.ev { line-height: 1.25; margin-bottom: 1px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
div.ev span.t { color: #555; }
tr { page-break-inside: avoid; }
${pageRule(o, true, 10)}
</style></head><body>
<h1>${escHtml(title)}</h1>
<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>
</body></html>`;
}

/** A message list as a table, Outlook's other print style.
 *
 *  Prints what is on screen, in the order it is on screen, so a filtered or
 *  sorted list prints as the thing you were looking at rather than as some
 *  other list built fresh. Every cell is escaped: these are subjects and
 *  sender names, which is to say text other people wrote. */
export function printableTableHtml(title: string, rows: readonly { from: string; subject: string; date: string; folder?: string }[], o?: PrintOptions): string {
	const cells = rows
		.map(
			(r) =>
				`<tr><td>${escHtml(r.from)}</td><td>${escHtml(r.subject)}</td><td class="nw">${escHtml(r.date)}</td><td>${escHtml(r.folder ?? "")}</td></tr>`
		)
		.join("");
	return `<!doctype html><html><head><meta charset="utf-8"><title>${escHtml(title)}</title><style>
:root { color-scheme: light; }
body { margin: 0; padding: 20px; background: #fff; color: #111;
  font-family: -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; font-size: ${o?.fontPt ?? 10}pt; }
h1 { font-size: 14pt; margin: 0 0 4px; }
p.count { margin: 0 0 14px; color: #555; font-size: 9pt; }
table { border-collapse: collapse; width: 100%; }
th { text-align: left; font-size: 9pt; text-transform: uppercase; letter-spacing: 0.03em; color: #555;
  border-bottom: 1.5px solid #999; padding: 4px 8px 4px 0; }
td { padding: 4px 8px 4px 0; border-bottom: 1px solid #ddd; vertical-align: top; }
td.nw { white-space: nowrap; }
/* a long table should not lose its headings on page two */
thead { display: table-header-group; }
tr { page-break-inside: avoid; }
${pageRule(o, false, 12)}
</style></head><body>
<h1>${escHtml(title)}</h1>
<p class="count">${rows.length} message${rows.length === 1 ? "" : "s"}</p>
<table><thead><tr><th>From</th><th>Subject</th><th>Received</th><th>Folder</th></tr></thead><tbody>${cells}</tbody></table>
</body></html>`;
}

/* ---------- signatures ---------- */

export interface MailSignature {
	id: string;
	name: string;
	html: string;
}

export interface SignatureUse {
	accountId: string;
	newId: string;
	replyId: string;
}

/** The signature to put on a message: the one this account uses for this
 *  kind of message, or nothing. An id that no longer names a signature
 *  resolves to nothing rather than to some other signature, because sending
 *  the wrong one is worse than sending none. */
export function signatureFor(sigs: readonly MailSignature[], use: readonly SignatureUse[], accountId: string, kind: "new" | "reply"): MailSignature | null {
	const u = use.find((x) => x.accountId === accountId);
	if (!u) return null;
	const id = kind === "new" ? u.newId : u.replyId;
	if (!id) return null;
	return sigs.find((s) => s.id === id) ?? null;
}

export interface InlineImage {
	cid: string;
	name: string;
	contentType: string;
	base64: string;
}

/** Pull data-url images out of outgoing HTML and point it at attachments.
 *
 *  A logo pasted into a signature lives as a data url while it is being
 *  edited, which is convenient and self-contained. Sent that way, though,
 *  Outlook and Gmail both refuse to render it: mail clients have blocked
 *  data-url images for years because they were an obvious tracking and
 *  payload dodge. The fix is what every mail client does, which is to send
 *  the bytes as an inline attachment and reference it by content id. */
export function extractInlineImages(html: string, idSeed: string): { html: string; images: InlineImage[] } {
	const images: InlineImage[] = [];
	let n = 0;
	const out = html.replace(/src\s*=\s*(["'])data:([^;]+);base64,([^"']+)\1/gi, (_all, quote: string, mime: string, b64: string) => {
		n++;
		const cid = `pd-${idSeed}-${n}`;
		const ext = mime.split("/")[1]?.replace(/[^a-z0-9]/gi, "") || "png";
		images.push({ cid, name: `image${n}.${ext}`, contentType: mime, base64: b64 });
		return `src=${quote}cid:${cid}${quote}`;
	});
	return { html: out, images };
}

/** Carry a single old signature into the named list, once.
 *
 *  Returns null when there is nothing to do, which is the common case: this
 *  runs on every load and must be silent unless it genuinely has an old
 *  signature and no new ones to convert. */
export function migrateSignature(oldHtml: string, sigs: readonly MailSignature[], accountIds: readonly string[], newId: string): { sigs: MailSignature[]; use: SignatureUse[] } | null {
	if (!oldHtml.trim() || sigs.length) return null;
	const sig = { id: newId, name: "Signature", html: oldHtml };
	return { sigs: [sig], use: accountIds.map((accountId) => ({ accountId, newId, replyId: newId })) };
}

/* ---------- new mail ---------- */

/** Which of a freshly synced list is worth announcing.
 *
 *  Three guards, and each exists because of a way this gets embarrassing:
 *  the baseline stops a first sync announcing a thousand old messages as if
 *  they just landed; the already-told set stops the same message arriving
 *  twice when a later sync returns it again; and read mail is never
 *  announced, because a message already read somewhere else is not news.
 *  Newest last, so the loudest notice is the most recent. */
export function newArrivals(messages: readonly PCMail[], baselineMs: number, told: ReadonlySet<string>, focusedOnly: boolean): PCMail[] {
	return messages
		.filter((m) => m.receivedMs > baselineMs && !told.has(m.id) && m.unread && (!focusedOnly || m.focused !== false))
		.sort((a, b) => a.receivedMs - b.receivedMs);
}

/** How a batch of arrivals reads in one line. */
export function arrivalSummary(arrivals: readonly PCMail[]): { title: string; body: string } {
	if (arrivals.length === 1) return { title: arrivals[0].from, body: arrivals[0].subject };
	const senders: string[] = [];
	for (const m of arrivals) if (!senders.includes(m.from)) senders.push(m.from);
	const who = senders.length <= 3 ? senders.join(", ") : `${senders.slice(0, 2).join(", ")} and ${senders.length - 2} others`;
	return { title: `${arrivals.length} new messages`, body: who };
}

/* ---------- the split inbox ---------- */

export type SectionKey = "priority" | "focused" | "notifications" | "other";

export interface MailSection {
	key: SectionKey;
	label: string;
	messages: PCMail[];
}

// prettier-ignore
const AUTOMATED_LOCAL = /^(no-?reply|do-?not-?reply|donotreply|notifications?|notify|alerts?|automated|automation|auto|system|sysadmin|mailer-daemon|postmaster|bounces?|jira|confluence|atlassian|github|gitlab|bitbucket|jenkins|builds?|ci|cd|monitoring|nagios|pagerduty|helpdesk|support-?bot|ticket|tickets)([.\-_+]|$)/i;

/** Whether an address is a machine rather than a person.
 *
 *  Only the local part is read, and only against names that are conventions
 *  rather than guesses: no-reply, notifications, the ticketing and build
 *  systems. A person at any of those domains still reads as a person, which
 *  matters because real colleagues mail from github.com and atlassian.net
 *  addresses all day. */
export function automatedSender(address: string): boolean {
	const local = (address ?? "").split("@")[0]?.trim();
	if (!local) return false;
	if (/no-?reply/i.test(local)) return true;
	return AUTOMATED_LOCAL.test(local);
}

/** Which section a message belongs in.
 *
 *  Three of the four tests read something Microsoft or the sender actually
 *  set rather than anything guessed here: the follow-up flag, the high
 *  importance mark, and Outlook's own Focused verdict, which is trained on
 *  how this mailbox actually behaves. Only Notifications is a heuristic, and
 *  it is a narrow one. A message with no verdict counts as focused, so mail
 *  cached before this shipped stays where the eye expects it. */
export function sectionOf(m: PCMail): SectionKey {
	if (m.flagged || m.priority) return "priority";
	if (automatedSender(m.fromAddress)) return "notifications";
	return m.focused === false ? "other" : "focused";
}

const SECTION_LABELS: Record<SectionKey, string> = {
	priority: "Priority",
	focused: "Focused",
	notifications: "Notifications",
	other: "Other",
};

/** A list split into its sections, in reading order, empty ones dropped.
 *
 *  Notifications sits above Other on purpose: ticket and build mail is work,
 *  and burying it under everything Outlook deprioritized would hide the part
 *  of the automated mail that actually needs answering. Order within a
 *  section is the order it arrived in, so whatever sorted the list still
 *  decides. */
export function splitSections(messages: PCMail[]): MailSection[] {
	const order: SectionKey[] = ["priority", "focused", "notifications", "other"];
	const buckets = new Map<SectionKey, PCMail[]>(order.map((k) => [k, []]));
	for (const m of messages) buckets.get(sectionOf(m))?.push(m);
	return order.filter((k) => (buckets.get(k) ?? []).length).map((k) => ({ key: k, label: SECTION_LABELS[k], messages: buckets.get(k) ?? [] }));
}

/* ---------- automatic replies ---------- */

/** Graph's dateTimeTimeZone with both halves present, which is what a
 *  setting written from here always has. The event mapper's own
 *  GraphDateTime is the same shape read loosely, since incoming events can
 *  be missing either half. */
export interface WhenSetting {
	dateTime: string;
	timeZone: string;
}

/** An instant as Graph's dateTimeTimeZone, always in UTC.
 *
 *  Exchange accepts a wall time paired with a zone name, but which zone
 *  spellings it takes is a swamp (Windows names, some IANA, not others), and
 *  getting it wrong means an auto-reply that starts at the wrong hour and
 *  says nothing about it. UTC is the one spelling nothing argues with, so
 *  the local time the user picked is converted here and sent absolute. */
export function toGraphDateTime(ms: number): WhenSetting {
	return { dateTime: new Date(ms).toISOString().slice(0, 19), timeZone: "UTC" };
}

/** Back the other way, for showing what the mailbox already has.
 *
 *  A UTC value is exact. Anything else is read as a wall time in whatever
 *  zone this machine is in, which is the best available guess and is right
 *  whenever the setting was made on a machine in the same zone. */
export function fromGraphDateTime(v: WhenSetting | undefined | null): number | null {
	const raw = (v?.dateTime ?? "").trim();
	if (!raw) return null;
	const zone = (v?.timeZone ?? "").trim().toUpperCase();
	// trim Graph's seven-digit fraction, which Date cannot read
	const base = raw.replace(/(\.\d{3})\d+$/, "$1").replace(/Z$/i, "");
	const ms = Date.parse(zone === "UTC" || zone === "GMT" ? `${base}Z` : base);
	return Number.isFinite(ms) ? ms : null;
}

/* ---------- categories ---------- */

/** Outlook's twenty-five category colors, in Graph's own order.
 *
 *  Graph hands back the name of a slot ("preset7") rather than a color, so
 *  the palette has to live somewhere, and it may as well be somewhere
 *  testable. The order is Outlook's: the light half first, then the dark. */
// prettier-ignore
const CATEGORY_PRESETS = [
	"#e74856", "#ff8c00", "#c19c00", "#ffd700", "#57a300", "#00b294", "#8cbd18", "#0078d4",
	"#8764b8", "#c30052", "#69797e", "#4a5459", "#7a7574", "#5d5a58", "#1f1f1f", "#a4262c",
	"#d13438", "#6e4b1e", "#986f0b", "#0b6a0b", "#005e50", "#4c6b0f", "#003966", "#5c2e91",
	"#93003f",
];

/** Outlook's own names for the twenty-five colors, so a color picker offers
 *  "Blue" rather than "preset7". */
// prettier-ignore
export const CATEGORY_COLOR_NAMES = [
	"Red", "Orange", "Brown", "Yellow", "Green", "Teal", "Olive", "Blue",
	"Purple", "Cranberry", "Steel", "Dark steel", "Gray", "Dark gray", "Black", "Dark red",
	"Dark orange", "Dark brown", "Dark yellow", "Dark green", "Dark teal", "Dark olive", "Dark blue", "Dark purple",
	"Dark cranberry",
];

/** The color behind a category, or a neutral gray when the mailbox gave a
 *  slot name nothing is known about. A category with no color set at all
 *  reads as "none" in Graph, which is a real state and not an error. */
export function categoryColor(preset: string): string {
	const m = /^preset(\d+)$/i.exec((preset ?? "").trim());
	if (!m) return "#8a8886";
	return CATEGORY_PRESETS[Number(m[1])] ?? "#8a8886";
}

/** Add or remove one category from a message's list, case-insensitively,
 *  keeping the order the mailbox had. Graph stores categories by their
 *  display name, so the comparison is on names and duplicates only differing
 *  in case would be two categories in Outlook's eyes as well. */
export function toggleCategory(current: string[] | undefined, name: string): string[] {
	const have = current ?? [];
	const hit = have.some((c) => c.toLowerCase() === name.toLowerCase());
	return hit ? have.filter((c) => c.toLowerCase() !== name.toLowerCase()) : [...have, name];
}

/** A category dressed as a folder id, so a category can be selected, pinned,
 *  and refreshed through the same paths a real folder uses.
 *
 *  The prefix cannot collide with a real id: Graph folder ids are base64url,
 *  which has no colon in it. The category's own name rides along because
 *  categories have no stable id on a message, only a display name. */
export const CATEGORY_PREFIX = "__cat__:";

export function categoryFolderId(name: string): string {
	return `${CATEGORY_PREFIX}${name}`;
}

/** The category a folder id stands for, or null when it names a real folder.
 *  An empty name is not a category: Graph refuses to store one, so a bare
 *  prefix can only be corruption, and reading it as "every message" would be
 *  the wrong guess. */
export function folderIdCategory(folderId: string): string | null {
	if (!folderId.startsWith(CATEGORY_PREFIX)) return null;
	const name = folderId.slice(CATEGORY_PREFIX.length);
	return name ? name : null;
}

/** Whether a message still carries the category its list was built from.
 *  Categories come off a message from inside the list that gathered them, so
 *  the view has to be able to drop a row the mailbox no longer agrees with. */
export function inCategory(m: PCMail, name: string): boolean {
	return (m.categories ?? []).some((c) => c.toLowerCase() === name.toLowerCase());
}

/* ---------- shortcuts ---------- */

export interface Shortcut {
	id: string;
	/** Which group it sits in; empty means the unnamed first group. */
	group: string;
	label: string;
	kind: "folder" | "search" | "note" | "url";
	/** A folder's account, for the folder kind only. */
	accountId?: string;
	/** The folder id, the query, the vault path, or the address. */
	target: string;
}

/** Shortcuts arranged into their groups, in the order they were made.
 *
 *  A group appears where its first member does, so moving a shortcut to the
 *  top moves its group there too, which is what someone rearranging a list
 *  by hand expects. The unnamed group leads, since that is where anything
 *  added without thinking about groups lands. */
export function groupShortcuts(list: readonly Shortcut[]): { name: string; items: Shortcut[] }[] {
	const out: { name: string; items: Shortcut[] }[] = [];
	for (const s of list) {
		const name = s.group.trim();
		const hit = out.find((g) => g.name === name);
		if (hit) hit.items.push(s);
		else out.push({ name, items: [s] });
	}
	return out.sort((a, b) => (a.name === "" ? -1 : b.name === "" ? 1 : 0));
}

/** What a shortcut is called when nothing was typed: the thing itself,
 *  shortened enough to sit in a list. */
export function defaultShortcutLabel(kind: Shortcut["kind"], target: string): string {
	if (kind === "note") return target.replace(/\.md$/i, "").split("/").pop() || target;
	if (kind === "url") return target.replace(/^https?:\/\//i, "").replace(/\/$/, "").slice(0, 60);
	if (kind === "search") return target.slice(0, 60);
	return target.slice(0, 40);
}

/* ---------- journal ---------- */

export interface JournalDay {
	key: string;
	meetings: { when: string; title: string; who: string }[];
	received: { when: string; who: string; subject: string }[];
	sent: { when: string; who: string; subject: string }[];
	notes: { title: string }[];
}

/** A day's activity, which is what Outlook's Journal was for before it was
 *  retired: what you sat in, who wrote to you, and what you sent.
 *
 *  Declined meetings are left out because you did not attend them, and an
 *  all-day event is listed without a time because it did not happen at one. */
export function buildJournal(
	key: string,
	events: readonly PCEvent[],
	received: readonly PCMail[],
	sent: readonly PCMail[],
	notes: readonly { title: string; changedMs: number }[],
	use24h: boolean
): JournalDay {
	const from = msOfKey(key);
	const to = from + 86400000;
	const inDay = (ms: number) => ms >= from && ms < to;
	const t = (ms: number) => fmtTimeOfMs(ms, use24h, true);
	return {
		key,
		meetings: events
			.filter((ev) => !ev.declined && eventsOnDay([ev], key).length)
			.sort((a, b) => a.startMs - b.startMs)
			.map((ev) => ({ when: ev.allDay ? "All day" : t(ev.startMs), title: ev.title, who: (ev.attendees ?? []).slice(0, 4).join(", ") })),
		received: received
			.filter((m) => inDay(m.receivedMs))
			.sort((a, b) => a.receivedMs - b.receivedMs)
			.map((m) => ({ when: t(m.receivedMs), who: m.from, subject: m.subject })),
		sent: sent
			.filter((m) => inDay(m.receivedMs))
			.sort((a, b) => a.receivedMs - b.receivedMs)
			.map((m) => ({ when: t(m.receivedMs), who: m.from, subject: m.subject })),
		notes: notes.filter((n) => inDay(n.changedMs)).map((n) => ({ title: n.title })),
	};
}

/** A day as markdown, for pasting into a daily note. Sections with nothing
 *  in them are left out rather than printed empty, since a heading with
 *  nothing under it is worse than no heading. */
export function journalMarkdown(d: JournalDay): string {
	const out: string[] = [`## ${fmtDayHeading(d.key)}`];
	const section = (title: string, lines: string[]) => {
		if (!lines.length) return;
		out.push("", `### ${title}`, ...lines);
	};
	section(
		"Meetings",
		d.meetings.map((m) => `- **${m.when}** ${m.title}${m.who ? ` with ${m.who}` : ""}`)
	);
	section(
		"Sent",
		d.sent.map((m) => `- **${m.when}** ${m.subject}`)
	);
	section(
		"Received",
		d.received.map((m) => `- **${m.when}** ${m.who}: ${m.subject}`)
	);
	section(
		"Notes",
		d.notes.map((n) => `- ${n.title}`)
	);
	if (out.length === 1) out.push("", "_Nothing recorded._");
	return out.join("\n") + "\n";
}

/* ---------- calendar search ---------- */

export interface EventQuery {
	words: string;
	title: string;
	people: string;
	location: string;
	/** A source key, or empty for every calendar. */
	calendar: string;
	onlineOnly: boolean;
	allDayOnly: boolean;
	withPeopleOnly: boolean;
}

/** Events matching a search, soonest first.
 *
 *  Every text field is an AND of its words, so "quarterly review" narrows
 *  rather than widening, and each is matched against the part of the event it
 *  names: `title` against the title alone, `people` against organizer and
 *  attendees, `words` against all of it including the description. Matching
 *  is case-blind, since nobody types a meeting's capitals from memory. */
export function matchEvents(events: readonly PCEvent[], q: EventQuery): PCEvent[] {
	const words = (s: string) =>
		s
			.toLowerCase()
			.split(/\s+/)
			.filter(Boolean);
	const all = words(q.words);
	const inTitle = words(q.title);
	const inWho = words(q.people);
	const inWhere = words(q.location);

	return events
		.filter((ev) => {
			if (q.calendar && ev.sourceId !== q.calendar) return false;
			if (q.onlineOnly && !ev.joinUrl) return false;
			if (q.allDayOnly && !ev.allDay) return false;
			const people = [ev.organizer ?? "", ...(ev.attendees ?? []), ...(ev.attendeeDetail ?? []).map((a) => `${a.name ?? ""} ${a.email ?? ""}`)].join(" ").toLowerCase();
			if (q.withPeopleOnly && !(ev.attendees ?? []).length) return false;
			const title = (ev.title ?? "").toLowerCase();
			const where = (ev.location ?? "").toLowerCase();
			const everything = `${title} ${where} ${people} ${(ev.description ?? "").toLowerCase()} ${(ev.calendarName ?? "").toLowerCase()}`;
			if (!all.every((w) => everything.includes(w))) return false;
			if (!inTitle.every((w) => title.includes(w))) return false;
			if (!inWho.every((w) => people.includes(w))) return false;
			if (!inWhere.every((w) => where.includes(w))) return false;
			return true;
		})
		.sort((a, b) => a.startMs - b.startMs);
}

/** Whether a search asks anything at all, so an empty dialog does not offer
 *  to list a year of calendar as if it were a result. */
export function eventQueryIsEmpty(q: EventQuery): boolean {
	return !q.words.trim() && !q.title.trim() && !q.people.trim() && !q.location.trim() && !q.calendar && !q.onlineOnly && !q.allDayOnly && !q.withPeopleOnly;
}

/* ---------- unsubscribe ---------- */

export interface UnsubscribeInfo {
	/** An https link the sender says accepts a one-click POST (RFC 8058).
	 *  Only set when the List-Unsubscribe-Post header vouches for it. */
	oneClickUrl?: string;
	/** An https link to open in a browser and finish by hand. */
	webUrl?: string;
	/** An address that unsubscribes you when mailed. */
	mailto?: { to: string; subject: string; body: string };
}

/** Read the unsubscribe headers a bulk sender is required to set.
 *
 *  List-Unsubscribe (RFC 2369) is a comma-separated list of URIs in angle
 *  brackets, any mix of mailto: and https:. List-Unsubscribe-Post (RFC 8058)
 *  is the sender promising the https one will accept a plain POST, which is
 *  what makes a one-click unsubscribe possible without opening a page.
 *
 *  Nothing here is trusted enough to act on by itself: these values are
 *  written by whoever sent the mail, so the caller shows the destination and
 *  asks before anything is opened, posted, or sent. */
export function parseUnsubscribe(headers: { name?: string; value?: string }[] | undefined): UnsubscribeInfo | null {
	if (!headers?.length) return null;
	const find = (want: string) => headers.find((h) => (h.name ?? "").toLowerCase() === want)?.value ?? "";
	const raw = find("list-unsubscribe");
	if (!raw.trim()) return null;
	const oneClick = /list-unsubscribe\s*=\s*one-click/i.test(find("list-unsubscribe-post"));
	const out: UnsubscribeInfo = {};
	for (const m of raw.matchAll(/<([^>]+)>/g)) {
		const uri = m[1].trim();
		if (/^https?:\/\//i.test(uri)) {
			if (!out.webUrl) out.webUrl = uri;
			if (oneClick && !out.oneClickUrl) out.oneClickUrl = uri;
		} else if (/^mailto:/i.test(uri) && !out.mailto) {
			const rest = uri.slice(7);
			const q = rest.indexOf("?");
			const to = (q < 0 ? rest : rest.slice(0, q)).trim();
			if (!to.includes("@")) continue;
			const params = new URLSearchParams(q < 0 ? "" : rest.slice(q + 1));
			out.mailto = { to, subject: params.get("subject") ?? "unsubscribe", body: params.get("body") ?? "unsubscribe" };
		}
	}
	return out.oneClickUrl || out.webUrl || out.mailto ? out : null;
}

/** How an unsubscribe will actually be carried out, in plain words, so the
 *  confirmation says what is about to happen rather than just "unsubscribe". */
export function unsubscribePlan(u: UnsubscribeInfo): { kind: "post" | "open" | "mail"; target: string } {
	if (u.oneClickUrl) return { kind: "post", target: u.oneClickUrl };
	if (u.webUrl) return { kind: "open", target: u.webUrl };
	return { kind: "mail", target: u.mailto?.to ?? "" };
}

/* ---------- inbox rules ---------- */

/** The parts of a rule this plugin offers to edit. Outlook can set a great
 *  deal more, which is exactly why this is a named subset rather than a
 *  pretence at the whole thing. */
export interface RuleEdit {
	name: string;
	enabled: boolean;
	fromContains: string;
	subjectContains: string;
	bodyContains: string;
	toContains: string;
	hasAttachments: boolean;
	highImportance: boolean;
	moveToFolderId: string;
	markAsRead: boolean;
	markImportance: "" | "low" | "normal" | "high";
	deleteIt: boolean;
	stopProcessing: boolean;
}

export const EMPTY_RULE: RuleEdit = {
	name: "",
	enabled: true,
	fromContains: "",
	subjectContains: "",
	bodyContains: "",
	toContains: "",
	hasAttachments: false,
	highImportance: false,
	moveToFolderId: "",
	markAsRead: false,
	markImportance: "",
	deleteIt: false,
	stopProcessing: false,
};

/** Comma-separated text to the string list Graph wants, and back. */
const listOf = (s: string): string[] =>
	s
		.split(/[,;]+/)
		.map((x) => x.trim())
		.filter(Boolean);
const textOf = (v: unknown): string => (Array.isArray(v) ? (v as unknown[]).map(String).join(", ") : "");

/** A stored rule read back into the editable subset. */
export function ruleToEdit(r: { displayName: string; isEnabled: boolean; conditions: Record<string, unknown>; actions: Record<string, unknown> }): RuleEdit {
	const c = r.conditions ?? {};
	const a = r.actions ?? {};
	return {
		name: r.displayName,
		enabled: r.isEnabled,
		fromContains: textOf(c.senderContains),
		subjectContains: textOf(c.subjectContains),
		bodyContains: textOf(c.bodyContains),
		toContains: textOf(c.recipientContains),
		hasAttachments: c.hasAttachments === true,
		highImportance: c.importance === "high",
		moveToFolderId: typeof a.moveToFolder === "string" ? a.moveToFolder : "",
		markAsRead: a.markAsRead === true,
		markImportance: a.markImportance === "low" || a.markImportance === "normal" || a.markImportance === "high" ? a.markImportance : "",
		deleteIt: a.delete === true,
		stopProcessing: a.stopProcessingRules === true,
	};
}

/** An edit folded back into a rule body for Graph.
 *
 *  The existing conditions and actions are carried through and only the
 *  managed keys are overwritten, because PATCH replaces these objects whole.
 *  Without the merge, editing a rule Outlook made would silently drop every
 *  condition this plugin does not model, which is the kind of data loss you
 *  would not notice until the mail stopped being filed. A managed field left
 *  empty deletes its key rather than writing an empty value, since Graph
 *  treats an empty list as a condition that matches nothing. */
export function ruleToBody(e: RuleEdit, existing?: { conditions: Record<string, unknown>; actions: Record<string, unknown> }): Record<string, unknown> {
	const conditions: Record<string, unknown> = { ...(existing?.conditions ?? {}) };
	const actions: Record<string, unknown> = { ...(existing?.actions ?? {}) };
	const set = (o: Record<string, unknown>, key: string, value: unknown, keep: boolean) => {
		if (keep) o[key] = value;
		else delete o[key];
	};
	set(conditions, "senderContains", listOf(e.fromContains), !!e.fromContains.trim());
	set(conditions, "subjectContains", listOf(e.subjectContains), !!e.subjectContains.trim());
	set(conditions, "bodyContains", listOf(e.bodyContains), !!e.bodyContains.trim());
	set(conditions, "recipientContains", listOf(e.toContains), !!e.toContains.trim());
	set(conditions, "hasAttachments", true, e.hasAttachments);
	set(conditions, "importance", "high", e.highImportance);
	set(actions, "moveToFolder", e.moveToFolderId, !!e.moveToFolderId);
	set(actions, "markAsRead", true, e.markAsRead);
	set(actions, "markImportance", e.markImportance, !!e.markImportance);
	set(actions, "delete", true, e.deleteIt);
	set(actions, "stopProcessingRules", true, e.stopProcessing);
	return { displayName: e.name.trim() || "Untitled rule", isEnabled: e.enabled, conditions, actions };
}

/** Whether a rule carries anything this plugin cannot show, so the editor
 *  can say so rather than letting someone believe they see all of it. */
export function ruleHasUnknownParts(r: { conditions: Record<string, unknown>; actions: Record<string, unknown>; exceptions: Record<string, unknown> }): boolean {
	const known = {
		conditions: ["senderContains", "subjectContains", "bodyContains", "recipientContains", "hasAttachments", "importance"],
		actions: ["moveToFolder", "markAsRead", "markImportance", "delete", "stopProcessingRules"],
	};
	if (Object.keys(r.exceptions ?? {}).length) return true;
	if (Object.keys(r.conditions ?? {}).some((k) => !known.conditions.includes(k))) return true;
	return Object.keys(r.actions ?? {}).some((k) => !known.actions.includes(k));
}

/** A rule as one readable line: what it looks for, then what it does. */
export function ruleSummary(e: RuleEdit, folderName?: string): string {
	const when: string[] = [];
	if (e.fromContains.trim()) when.push(`from ${e.fromContains.trim()}`);
	if (e.subjectContains.trim()) when.push(`subject has ${e.subjectContains.trim()}`);
	if (e.bodyContains.trim()) when.push(`body has ${e.bodyContains.trim()}`);
	if (e.toContains.trim()) when.push(`sent to ${e.toContains.trim()}`);
	if (e.hasAttachments) when.push("has an attachment");
	if (e.highImportance) when.push("is high importance");
	const then: string[] = [];
	if (e.moveToFolderId) then.push(`move to ${folderName || "a folder"}`);
	if (e.deleteIt) then.push("delete");
	if (e.markAsRead) then.push("mark read");
	if (e.markImportance) then.push(`mark ${e.markImportance} importance`);
	if (e.stopProcessing) then.push("stop processing rules");
	if (!when.length && !then.length) return "Does nothing yet";
	if (!when.length) return `Everything: ${then.join(", ")}`;
	if (!then.length) return `When ${when.join(" and ")}: nothing yet`;
	return `When ${when.join(" and ")}: ${then.join(", ")}`;
}

/* ---------- local search index ---------- */

/** Words worth indexing, lowercased. Splits on everything that is not a
 *  letter, a digit, or the few marks that hold an address or a ticket number
 *  together, so "ACME-510902" and "dana.lee@example.com" survive as findable
 *  units as well as breaking into their parts. */
export function tokenize(text: string): string[] {
	const out: string[] = [];
	for (const raw of (text ?? "").toLowerCase().split(/[^a-z0-9._@+-]+/)) {
		const t = raw.replace(/^[._@+-]+|[._@+-]+$/g, "");
		if (!t) continue;
		out.push(t);
		// an address or a hyphenated id is also findable by its pieces, so
		// "lee" finds dana.lee@example.com and "510902" finds the ticket
		if (/[._@+-]/.test(t)) for (const p of t.split(/[._@+-]+/)) if (p && p !== t) out.push(p);
	}
	return out;
}

export interface IndexDoc {
	id: string;
	subject: string;
	from: string;
	body: string;
	ms: number;
	unread: boolean;
	flagged: boolean;
	hasAttachments: boolean;
}

export interface MailIndex {
	/** token to the ids that carry it. */
	postings: Map<string, Set<string>>;
	docs: Map<string, IndexDoc>;
}

/** Build an inverted index over what is already cached. Subject and sender
 *  are indexed always; the body only for messages whose body has been read,
 *  which is why local search finds more the longer a mailbox is used. */
export function buildIndex(docs: IndexDoc[]): MailIndex {
	const postings = new Map<string, Set<string>>();
	const byId = new Map<string, IndexDoc>();
	for (const d of docs) {
		byId.set(d.id, d);
		for (const t of new Set([...tokenize(d.subject), ...tokenize(d.from), ...tokenize(d.body)])) {
			let set = postings.get(t);
			if (!set) postings.set(t, (set = new Set()));
			set.add(d.id);
		}
	}
	return { postings, docs: byId };
}

export interface ParsedQuery {
	/** Bare words, matched as prefixes against the index. */
	terms: string[];
	/** Quoted runs, matched as substrings of the whole document. */
	phrases: string[];
	from: string[];
	subject: string[];
	unread?: boolean;
	flagged?: boolean;
	attachments?: boolean;
	/** Local midnight bounds from after: and before:, as epoch ms. `before`
	 *  is the end of that day, since "before Friday" said out loud does not
	 *  mean "before midnight on Friday morning". */
	afterMs?: number;
	beforeMs?: number;
}

/** A date written the way a search box takes it: a day key, or nothing. */
function parseDayBound(v: string, endOfDay: boolean): number | undefined {
	const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v.trim());
	if (!m) return undefined;
	const ms = msOfKey(`${m[1]}-${m[2]}-${m[3]}`);
	return endOfDay ? ms + 86399999 : ms;
}

/** The small query language a mail search box is expected to speak:
 *  `from:steve`, `subject:invoice`, `is:unread`, `is:flagged`,
 *  `has:attachment`, "a quoted phrase", and bare words for everything else. */
export function parseQuery(q: string): ParsedQuery {
	const out: ParsedQuery = { terms: [], phrases: [], from: [], subject: [] };
	for (const m of (q ?? "").matchAll(/"([^"]*)"|(\S+)/g)) {
		const phrase = m[1];
		if (phrase !== undefined) {
			if (phrase.trim()) out.phrases.push(phrase.trim().toLowerCase());
			continue;
		}
		const tok = m[2];
		const pair = tok.match(/^(from|subject|is|has|after|before):(.+)$/i);
		if (!pair) {
			out.terms.push(...tokenize(tok));
			continue;
		}
		const key = pair[1].toLowerCase();
		const val = pair[2].toLowerCase();
		if (key === "from") out.from.push(val);
		else if (key === "subject") out.subject.push(val);
		else if (key === "is") {
			if (val === "unread") out.unread = true;
			else if (val === "read") out.unread = false;
			else if (val === "flagged") out.flagged = true;
		} else if (key === "has") {
			if (val.startsWith("attach")) out.attachments = true;
		} else if (key === "after") out.afterMs = parseDayBound(val, false);
		else if (key === "before") out.beforeMs = parseDayBound(val, true);
	}
	return out;
}

export interface SearchFields {
	words: string;
	from: string;
	subject: string;
	phrase: string;
	unread: boolean;
	flagged: boolean;
	attachments: boolean;
	after: string;
	before: string;
}

/** Fields from a search dialog into the query language itself.
 *
 *  The dialog composes the same text you could type, rather than a private
 *  structure: what it builds is shown back to you, so the box and the boxes
 *  are the same search and using one teaches the other. */
export function buildQuery(f: Partial<SearchFields>): string {
	const parts: string[] = [];
	const words = (f.words ?? "").trim();
	if (words) parts.push(words);
	if ((f.phrase ?? "").trim()) parts.push(`"${(f.phrase ?? "").trim().replace(/"/g, "")}"`);
	for (const v of (f.from ?? "").split(/[,;]+/).map((x) => x.trim()).filter(Boolean)) parts.push(`from:${v}`);
	for (const v of (f.subject ?? "").split(/\s+/).filter(Boolean)) parts.push(`subject:${v}`);
	if (f.unread) parts.push("is:unread");
	if (f.flagged) parts.push("is:flagged");
	if (f.attachments) parts.push("has:attachment");
	if ((f.after ?? "").trim()) parts.push(`after:${(f.after ?? "").trim()}`);
	if ((f.before ?? "").trim()) parts.push(`before:${(f.before ?? "").trim()}`);
	return parts.join(" ");
}

/** The part of a query the mailbox can answer, and nothing else.
 *
 *  Graph's search speaks words, phrases, from: and subject:, and treats
 *  anything else as text to look for: send it `is:unread` and it hunts for
 *  messages containing the string "is:unread", which is worse than useless
 *  because it looks like it worked. The rest is applied here instead. */
export function graphSearchText(p: ParsedQuery): string {
	const parts: string[] = [...p.terms];
	for (const x of p.phrases) parts.push(`"${x}"`);
	for (const x of p.from) parts.push(`from:${x}`);
	for (const x of p.subject) parts.push(`subject:${x}`);
	return parts.join(" ");
}

/** Whether a message passes the parts of a query the mailbox did not apply.
 *  Used on results that came back from the server, which knew nothing about
 *  the flags or the dates. */
export function passesLocalFilters(m: PCMail, p: ParsedQuery): boolean {
	if (p.unread !== undefined && m.unread !== p.unread) return false;
	if (p.flagged && !m.flagged) return false;
	if (p.attachments && !m.hasAttachments) return false;
	if (p.afterMs !== undefined && m.receivedMs < p.afterMs) return false;
	if (p.beforeMs !== undefined && m.receivedMs > p.beforeMs) return false;
	return true;
}

/** Ids that satisfy a parsed query, best first.
 *
 *  Bare words match as prefixes, so results narrow with every keystroke
 *  rather than appearing only once a word is finished. Every word has to
 *  match something (AND, the way a mail search is expected to behave), and a
 *  hit in the subject or the sender outranks one buried in a body. */
export function searchIndex(index: MailIndex, q: string, limit = 100): string[] {
	const p = parseQuery(q);
	const anything =
		p.terms.length || p.phrases.length || p.from.length || p.subject.length || p.unread !== undefined || p.flagged || p.attachments || p.afterMs !== undefined || p.beforeMs !== undefined;
	if (!anything) return [];

	let candidates: Set<string> | null = null;
	for (const term of p.terms) {
		const hits = new Set<string>();
		// an exact token is the common case; the prefix sweep covers the word
		// still being typed
		for (const id of index.postings.get(term) ?? []) hits.add(id);
		for (const [tok, ids] of index.postings) if (tok.length > term.length && tok.startsWith(term)) for (const id of ids) hits.add(id);
		candidates = candidates ? new Set([...candidates].filter((id) => hits.has(id))) : hits;
		if (!candidates.size) return [];
	}
	const pool = candidates ? [...candidates] : [...index.docs.keys()];

	const scored: { id: string; rank: number; ms: number }[] = [];
	for (const id of pool) {
		const d = index.docs.get(id);
		if (!d) continue;
		const subject = d.subject.toLowerCase();
		const from = d.from.toLowerCase();
		const body = d.body.toLowerCase();
		if (p.afterMs !== undefined && d.ms < p.afterMs) continue;
		if (p.beforeMs !== undefined && d.ms > p.beforeMs) continue;
		if (p.unread !== undefined && d.unread !== p.unread) continue;
		if (p.flagged && !d.flagged) continue;
		if (p.attachments && !d.hasAttachments) continue;
		if (!p.from.every((x) => from.includes(x))) continue;
		if (!p.subject.every((x) => subject.includes(x))) continue;
		if (!p.phrases.every((x) => subject.includes(x) || from.includes(x) || body.includes(x))) continue;
		const strong = p.terms.some((t) => subject.includes(t) || from.includes(t)) || p.phrases.some((x) => subject.includes(x) || from.includes(x));
		scored.push({ id, rank: strong ? 0 : 1, ms: d.ms });
	}
	return scored
		.sort((a, b) => a.rank - b.rank || b.ms - a.ms)
		.slice(0, limit)
		.map((s) => s.id);
}

/* ---------- recipient autocomplete ---------- */

export interface ContactHit {
	name: string;
	email: string;
	/** How many times this address has been seen, which is the whole of the
	 *  ranking signal along with recency. */
	count: number;
	lastMs: number;
}

/** Fold every sighting of an address into one ranked entry each.
 *
 *  The name is whichever sighting was most recent and actually carried one,
 *  since people change how they sign and the latest is the one they answer
 *  to. Ranking is count first and recency second: the address you write to
 *  constantly should lead even if you wrote to someone else this morning. */
export function rankContacts(seen: { name?: string; email: string; ms: number }[]): ContactHit[] {
	const byEmail = new Map<string, ContactHit>();
	for (const s of seen) {
		const email = s.email.trim().toLowerCase();
		if (!email.includes("@") || email.length < 3) continue;
		const name = (s.name ?? "").trim();
		const hit = byEmail.get(email);
		if (!hit) {
			byEmail.set(email, { name, email, count: 1, lastMs: s.ms });
			continue;
		}
		hit.count++;
		if (s.ms > hit.lastMs) {
			hit.lastMs = s.ms;
			if (name) hit.name = name;
		} else if (!hit.name && name) hit.name = name;
	}
	return [...byEmail.values()].sort((a, b) => b.count - a.count || b.lastMs - a.lastMs);
}

/** The suggestions for what has been typed so far.
 *
 *  A prefix beats a match in the middle, because someone typing "ste" means
 *  Steve rather than "webmaster@stevecorp". Both the address and every word
 *  of the name are candidates for the prefix, so "palm" finds Steve Palm and
 *  "steve.p" finds the address. */
export function matchContacts(index: ContactHit[], query: string, limit = 8): ContactHit[] {
	const q = query.trim().toLowerCase();
	if (!q) return index.slice(0, limit);
	const words = (c: ContactHit) => [c.email, c.email.split("@")[0], ...c.name.toLowerCase().split(/[\s,._-]+/)].filter(Boolean);
	const scored: { c: ContactHit; rank: number }[] = [];
	for (const c of index) {
		const w = words(c);
		if (w.some((x) => x.startsWith(q))) scored.push({ c, rank: 0 });
		else if (c.email.includes(q) || c.name.toLowerCase().includes(q)) scored.push({ c, rank: 1 });
	}
	return scored
		.sort((a, b) => a.rank - b.rank || b.c.count - a.c.count || b.c.lastMs - a.c.lastMs)
		.slice(0, limit)
		.map((s) => s.c);
}

export interface SavedContact {
	name: string;
	email: string;
	company?: string;
	title?: string;
	phone?: string;
}

export interface PersonCard extends ContactHit {
	/** In the mailbox's own Contacts, rather than only inferred from mail. */
	saved: boolean;
	company?: string;
	title?: string;
	phone?: string;
}

/** The address book and the correspondence, as one list.
 *
 *  A saved contact you have never written to still belongs here, and a
 *  colleague you write to daily who was never saved belongs here more. So
 *  the two merge by address, keeping the saved details and the real counts,
 *  and the order is how much you actually deal with someone first, then the
 *  saved names you have not, alphabetically. A contact with several
 *  addresses appears once per address, because the address is what you send
 *  to and the one you actually use is the one that will rise. */
export function mergePeople(seen: readonly ContactHit[], contacts: readonly SavedContact[]): PersonCard[] {
	const by = new Map<string, PersonCard>();
	for (const c of seen) by.set(c.email, { ...c, saved: false });
	for (const c of contacts) {
		const email = c.email.trim().toLowerCase();
		if (!email.includes("@")) continue;
		const hit = by.get(email);
		if (hit) {
			hit.saved = true;
			// the address book's name is the deliberate one; a name scraped
			// off a From line is whatever that sender happened to set
			if (c.name.trim()) hit.name = c.name.trim();
			hit.company = c.company || hit.company;
			hit.title = c.title || hit.title;
			hit.phone = c.phone || hit.phone;
		} else {
			by.set(email, { name: c.name.trim() || email, email, count: 0, lastMs: 0, saved: true, company: c.company, title: c.title, phone: c.phone });
		}
	}
	return [...by.values()].sort((a, b) => b.count - a.count || b.lastMs - a.lastMs || a.name.localeCompare(b.name));
}

/** The fragment being typed in a comma-separated recipient box: where it
 *  starts and what it says. Everything before it is already-entered
 *  addresses and must survive untouched when a suggestion is taken. */
export function currentAddressFragment(value: string, caret: number): { start: number; text: string } {
	const upto = value.slice(0, caret);
	const cut = Math.max(upto.lastIndexOf(","), upto.lastIndexOf(";"));
	const start = cut + 1;
	return { start, text: value.slice(start, caret).trim() };
}

/** A chosen address folded back into the box, replacing only the fragment
 *  that was being typed and leaving a trailing comma to carry on from. */
export function applyAddressChoice(value: string, caret: number, email: string): { value: string; caret: number } {
	const { start } = currentAddressFragment(value, caret);
	const before = value.slice(0, start);
	const after = value.slice(caret);
	const lead = before && !/[\s]$/.test(before) ? " " : "";
	const next = `${before}${lead}${email}, `;
	return { value: `${next}${after.replace(/^[\s,;]+/, "")}`, caret: next.length };
}

export interface WhenPreset {
	label: string;
	ms: number;
}

/** A chosen time written the way a confirmation should read it: the day in
 *  words when it is close, the date when it is not. */
export function fmtWhen(ms: number, use24h: boolean): string {
	const key = keyOfMs(ms);
	const today = keyOfDate(new Date());
	const days = dayDiff(today, key);
	const time = fmtTimeOfMs(ms, use24h, true);
	if (days === 0) return `today at ${time}`;
	if (days === 1) return `tomorrow at ${time}`;
	if (days > 1 && days < 7) return `${fmtDayHeading(key)} at ${time}`;
	return `${fmtDayShort(key, true)} at ${time}`;
}

/** The "later" menu every mail client offers, resolved against a local
 *  clock. Snooze uses it to decide when a message comes back and schedule
 *  send to decide when a draft goes out, so the two read the same.
 *
 *  Presets that have already passed today are left out rather than offered
 *  as a time in the past: at 9pm "This evening" is not a choice, and a menu
 *  that offers it is a menu that will quietly do nothing. */
export function whenPresets(nowMs: number): WhenPreset[] {
	const key = keyOfMs(nowMs);
	const at = (dayKey: string, hour: number) => msOfKey(dayKey) + hour * 3600000;
	const mins = minutesOfMs(nowMs);
	const out: WhenPreset[] = [];
	if (mins < 15 * 60) out.push({ label: "Later today", ms: nowMs + 3 * 3600000 });
	if (mins < 17 * 60) out.push({ label: "This evening", ms: at(key, 18) });
	out.push({ label: "Tomorrow morning", ms: at(addDays(key, 1), 8) });
	out.push({ label: "Tomorrow afternoon", ms: at(addDays(key, 1), 13) });
	// 0 is Sunday. The weekend means the coming Saturday, and asking for it
	// on a Saturday means the next one, not this morning.
	const dow = dayOfWeek(key);
	const toSat = dow === 6 ? 7 : dow === 0 ? 6 : 6 - dow;
	out.push({ label: "This weekend", ms: at(addDays(key, toSat), 8) });
	const toMon = dow === 1 ? 7 : (8 - dow) % 7;
	out.push({ label: "Next week", ms: at(addDays(key, toMon), 8) });
	return out;
}

/** A back-and-forth collapsed into one row: the messages newest first, and
 *  the marks and counts rolled up off all of them. */
export interface MailThread {
	/** Account and conversation together, since a conversation id is only
	 *  unique inside its own mailbox and two accounts must never braid. */
	key: string;
	/** Newest first, which is the order a collapsed thread expands into. */
	messages: PCMail[];
	/** The newest message: what a collapsed row shows, and what clicking it
	 *  opens. Outlook lands on the newest too. */
	latest: PCMail;
	unread: number;
	/** Everyone who wrote, newest first, for the collapsed row's name line. */
	senders: string[];
	hasAttachments: boolean;
	flagged: boolean;
	priority: boolean;
}

/** Group a rendered list into conversations, keeping the list's own order:
 *  a thread sits where its first message sat, so whatever sort produced the
 *  list still decides what comes first. A message with no conversation id is
 *  a thread of one, which is what makes this safe over a cache written
 *  before threading existed. */
export function groupThreads(messages: PCMail[]): MailThread[] {
	const byKey = new Map<string, PCMail[]>();
	const order: string[] = [];
	for (const m of messages) {
		const key = m.conversationId ? `${m.accountId}\u0000${m.conversationId}` : `\u0000id\u0000${m.id}`;
		const bucket = byKey.get(key);
		if (bucket) bucket.push(m);
		else {
			byKey.set(key, [m]);
			order.push(key);
		}
	}
	return order.map((key) => {
		const msgs = [...(byKey.get(key) ?? [])].sort((a, b) => b.receivedMs - a.receivedMs);
		const senders: string[] = [];
		for (const m of msgs) if (!senders.includes(m.from)) senders.push(m.from);
		return {
			key,
			messages: msgs,
			latest: msgs[0],
			unread: msgs.filter((m) => m.unread).length,
			senders,
			hasAttachments: msgs.some((m) => !!m.hasAttachments),
			flagged: msgs.some((m) => !!m.flagged),
			priority: msgs.some((m) => !!m.priority),
		};
	});
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
export function orderFolderTree(
	folders: MailFolder[],
	inboxId: string | null,
	collapsed?: Set<string>,
	/** Folder ids the user has dragged into an order of their own. Anything
	 *  named here leads, in this order; everything else keeps the default
	 *  alphabetical arrangement behind it, so ordering three folders by hand
	 *  does not scramble the other thirty. */
	custom?: string[]
): { folder: MailFolder; depth: number; expandable: boolean }[] {
	const ids = new Set(folders.map((f) => f.id));
	const byParent = new Map<string | null, MailFolder[]>();
	for (const f of folders) {
		const p = f.parentId && ids.has(f.parentId) ? f.parentId : null;
		const arr = byParent.get(p);
		if (arr) arr.push(f);
		else byParent.set(p, [f]);
	}
	const rank = new Map((custom ?? []).map((id, i) => [id, i]));
	const alpha = (a: MailFolder, b: MailFolder) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" });
	/** Custom placement first, then the old rule for everything untouched. */
	const order = (a: MailFolder, b: MailFolder) => {
		const ra = rank.get(a.id);
		const rb = rank.get(b.id);
		if (ra !== undefined && rb !== undefined) return ra - rb;
		if (ra !== undefined) return -1;
		if (rb !== undefined) return 1;
		return alpha(a, b);
	};
	const top = (byParent.get(null) ?? []).slice().sort((a, b) => {
		// the inbox stays on top unless it has been dragged somewhere itself
		const pinned = (f: MailFolder) => (f.id === inboxId && !rank.has(f.id) ? 0 : 1);
		return pinned(a) - pinned(b) || order(a, b);
	});
	const out: { folder: MailFolder; depth: number; expandable: boolean }[] = [];
	const walk = (f: MailFolder, depth: number) => {
		const kids = (byParent.get(f.id) ?? []).slice().sort(order);
		out.push({ folder: f, depth, expandable: kids.length > 0 });
		if (collapsed?.has(f.id)) return;
		for (const child of kids) walk(child, depth + 1);
	};
	for (const f of top) walk(f, 0);
	return out;
}

// prettier-ignore
const SYSTEM_FOLDERS = new Set([
	"inbox", "drafts", "sent items", "deleted items", "junk email", "outbox", "archive",
	"conversation history", "notes", "journal", "tasks", "rss feeds", "rss subscriptions",
	"sync issues", "conflicts", "local failures", "server failures", "clutter", "scheduled",
]);

/** Whether a folder is one the mailbox runs on rather than one you made.
 *
 *  Renaming Sent Items or deleting Drafts is not something to offer, and
 *  Graph's refusal would arrive as a raw error after the fact. The inbox is
 *  matched by id because it is the one folder whose id is already known for
 *  certain; the rest go by name, which is a good guard in an English mailbox
 *  and a harmless one in any other, since Graph still refuses. */
export function isSystemFolder(name: string, folderId: string, inboxId: string | null): boolean {
	if (inboxId && folderId === inboxId) return true;
	return SYSTEM_FOLDERS.has((name ?? "").trim().toLowerCase());
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

/* ---------- a message as Markdown ---------- */

/** One picture a body shows inline: the content id it points at with
 *  src="cid:…", or the data URL it carries in the src itself. `key` is what
 *  the Markdown pass looks the saved file up by, so both halves have to
 *  derive it the same way, which is why they share one walk of the HTML. */
export type MailInlineImage = { key: string; cid?: string; dataUrl?: string };

/** A content id as it can be compared: Graph reports it bare, some senders
 *  write it inside angle brackets, and the case never matters. */
export function normalizeCid(cid: string): string {
	return cid.trim().replace(/^<+|>+$/g, "").toLowerCase();
}

const attrOf = (tag: string, name: string): string =>
	new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s"'>]+))`, "i")
		.exec(tag)
		?.slice(1)
		.find((v) => v !== undefined) ?? "";

/** CSS lengths in pixels each, at the 96dpi every mail client renders at.
 *  Word writes an inserted picture's size in inches, so this is what turns
 *  Outlook's own "width:6.5in" into the 624 the message shows. */
const CSS_UNIT_PX: Record<string, number> = { px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, pt: 96 / 72, pc: 16 };

/** A tag's width or height in pixels, from the style or the attribute, or
 *  null when it says nothing definite (a percentage sizes to the message,
 *  and says nothing about how big the picture is). */
function pixelSize(tag: string, name: "width" | "height"): number | null {
	const styled = new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([\\d.]+)\\s*(px|in|cm|mm|pt|pc)`, "i").exec(attrOf(tag, "style"));
	const attr = /^([\d.]+)(px)?$/i.exec(attrOf(tag, name).trim());
	const [raw, unit] = styled ? [styled[1], styled[2].toLowerCase()] : attr ? [attr[1], "px"] : ["", ""];
	const n = parseFloat(raw) * (CSS_UNIT_PX[unit] ?? 0);
	return Number.isFinite(n) && n > 0 ? n : null;
}

/** The width in pixels a message asks a picture to be shown at, rounded to
 *  what an embed can carry. Null when the message leaves it to the reader. */
export function declaredWidth(tag: string): number | null {
	const w = pixelSize(tag, "width");
	return w === null ? null : Math.max(1, Math.round(w));
}

/** An embed with a pixel width written into it, in whichever link syntax the
 *  vault uses: ![[file|600]] or ![alt|600](file). */
export function withEmbedWidth(embed: string, width: number | null): string {
	if (!width) return embed;
	const wiki = /^!\[\[([^\]]+)\]\]$/.exec(embed.trim());
	if (wiki) return `![[${wiki[1]}|${width}]]`;
	const md = /^!\[([^\]]*)\]\(([\s\S]*)\)$/.exec(embed.trim());
	if (md) return `![${md[1] ? `${md[1]}|` : ""}${width}](${md[2]})`;
	return embed;
}

/** A picture's own size, read from the front of the file: PNG, GIF, BMP and
 *  JPEG, which covers what a message embeds. Null for anything else, and for
 *  SVG, which carries no pixel size at all. */
export function imageSize(bytes: Uint8Array): { w: number; h: number } | null {
	if (bytes.length < 24) return null;
	const be16 = (i: number) => (bytes[i] << 8) | bytes[i + 1];
	const be32 = (i: number) => ((bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3]) >>> 0;
	const le16 = (i: number) => bytes[i] | (bytes[i + 1] << 8);
	const le32 = (i: number) => (bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24)) >>> 0;
	if (be32(0) === 0x89504e47) return { w: be32(16), h: be32(20) }; // PNG: the IHDR chunk sits at a fixed offset
	if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return { w: le16(6), h: le16(8) }; // GIF
	if (bytes[0] === 0x42 && bytes[1] === 0x4d) return { w: le32(18), h: le32(22) }; // BMP
	if (be16(0) !== 0xffd8) return null;
	// JPEG keeps its size in the frame header, which is only reachable by
	// walking the segments: they vary in length and any number of them can
	// come first
	for (let i = 2; i + 9 < bytes.length; ) {
		if (bytes[i] !== 0xff) {
			i++;
			continue;
		}
		const marker = bytes[i + 1];
		// standalone markers carry no length to skip by
		if (marker === 0x01 || marker === 0xd8 || marker === 0xd9 || (marker >= 0xd0 && marker <= 0xd7)) {
			i += 2;
			continue;
		}
		if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return { w: be16(i + 7), h: be16(i + 5) };
		const len = be16(i + 2);
		if (len < 2) return null;
		i += 2 + len;
	}
	return null;
}

/** True for the invisible pixel a newsletter loads to report that a message
 *  was opened. Nothing is gained by writing those into a vault. */
function isTrackingPixel(tag: string): boolean {
	const w = pixelSize(tag, "width");
	const h = pixelSize(tag, "height");
	return (w !== null && w <= 2) || (h !== null && h <= 2);
}

/** Every <img> in a body, in order, with what it points at. */
function imagesIn(html: string): { tag: string; start: number; end: number; ref: MailInlineImage | null; tracking: boolean }[] {
	const out: { tag: string; start: number; end: number; ref: MailInlineImage | null; tracking: boolean }[] = [];
	let data = 0;
	for (const m of html.matchAll(/<img\b[^>]*>/gi)) {
		const tag = m[0];
		const src = attrOf(tag, "src").trim();
		const cid = /^cid:/i.test(src) ? normalizeCid(src.slice(4)) : "";
		// a data URL carries no id of its own, so its place in the message is
		// its name; both passes number them off the same walk
		const ref: MailInlineImage | null = cid ? { key: cid, cid } : /^data:image\//i.test(src) ? { key: `data:${data++}`, dataUrl: src } : null;
		out.push({ tag, start: m.index, end: m.index + tag.length, ref, tracking: isTrackingPixel(tag) });
	}
	return out;
}

/** The inline pictures a body shows that have to be saved before a note can
 *  show them: content-id references and data URLs, tracking pixels left
 *  behind. Remote images are not here, since a note can point at those the
 *  way the message does. */
export function mailInlineImages(html: string): MailInlineImage[] {
	const out: MailInlineImage[] = [];
	for (const img of imagesIn(html)) {
		if (!img.ref || img.tracking) continue;
		if (!out.some((r) => r.key === img.ref?.key)) out.push(img.ref);
	}
	return out;
}

/** The word a saved picture rides through the conversion as. Letters and
 *  digits only, because anything with punctuation in it comes back escaped. */
const imgToken = (n: number) => `PDMAILIMAGE${n}ENDIMAGE`;

/** A saved picture, ready to embed: the link the vault's own link style
 *  wrote, and the size of the file itself, which decides how wide the embed
 *  is written when the message never said. */
export type MailEmbed = { link: string; naturalWidth?: number };

/** How wide a picture the message did not size is written at. A little under
 *  the readable column, so a photo lands as a picture in a note rather than
 *  as a wall the text below it starts under. */
const EMBED_MAX_WIDTH = 600;

/** The bytes and file extension a data URL carries, or null when it is not
 *  base64 data. */
export function parseDataUrl(url: string): { base64: string; ext: string } | null {
	const m = /^data:([\w.+-]+\/[\w.+-]+)?[^,]*;base64,([\s\S]+)$/i.exec(url.trim());
	if (!m) return null;
	return { base64: m[2].replace(/\s+/g, ""), ext: imageExtension("", m[1] ?? "") };
}

/** What to call a saved picture's file. The attachment's own name usually
 *  carries the extension; the content type answers when it does not. */
export function imageExtension(name: string, contentType: string): string {
	const named = /\.([a-z0-9]{2,5})$/i.exec(name.trim())?.[1];
	if (named) return named.toLowerCase();
	const sub = /^image\/([\w.+-]+)/i.exec(contentType.trim())?.[1]?.toLowerCase() ?? "";
	if (sub === "jpeg") return "jpg";
	if (sub === "svg+xml") return "svg";
	if (sub === "x-icon" || sub === "vnd.microsoft.icon") return "ico";
	return /^[a-z0-9]+$/.test(sub) ? sub : "png";
}

/** A message body as Markdown for a note.
 *
 *  The stripped text a saved message used to carry threw away every heading,
 *  list, link and picture in it. Obsidian's own htmlToMarkdown does the
 *  conversion, passed in so this file stays free of the API, with the
 *  mail-specific work either side: styles, scripts and tracking pixels out
 *  first, and each inline picture swapped for the embed of the file already
 *  written beside the note. An embed rides through the conversion as a plain
 *  word rather than as a link, because a link written into the HTML comes
 *  back out with its brackets escaped. A picture nothing was saved for is
 *  dropped rather than left pointing at a cid: nothing can open.
 *
 *  Each embed is written at the size the message shows the picture at, which
 *  is what keeps a photo inserted at 6.5 inches from arriving as the 1237
 *  pixels the camera took. When the message leaves the size to the reader, a
 *  picture bigger than a page gets one, since a note is not a mail client and
 *  has no width of its own to shrink it to. */
export function mailBodyMarkdown(html: string, toMd: (html: string) => string, embeds: ReadonlyMap<string, MailEmbed> = new Map()): string {
	const used: string[] = [];
	let stripped = "";
	let last = 0;
	for (const img of imagesIn(html)) {
		stripped += html.slice(last, img.start);
		last = img.end;
		if (img.tracking) continue;
		const embed = img.ref ? embeds.get(img.ref.key) : undefined;
		if (embed) {
			const natural = embed.naturalWidth ?? 0;
			stripped += ` ${imgToken(used.length)} `;
			used.push(withEmbedWidth(embed.link, declaredWidth(img.tag) ?? (natural > EMBED_MAX_WIDTH ? EMBED_MAX_WIDTH : null)));
		} else if (!img.ref) {
			stripped += img.tag; // remote: the note points at it the way the message does
		}
	}
	stripped += html.slice(last);
	const md = toMd(
		stripped
			.replace(/<style[\s\S]*?<\/style>/gi, "")
			.replace(/<script[\s\S]*?<\/script>/gi, "")
			.replace(/<!--[\s\S]*?-->/g, "")
			.replace(/<\/?o:p[^>]*>/gi, "")
	);
	return used
		.reduce(
			(s, embed, i) => s.split(imgToken(i)).join(embed),
			// a linked picture converts to [token](url), and an embed cannot sit
			// inside a link, so the picture keeps its place and the link goes
			md
				.replace(/\[\s*(PDMAILIMAGE\d+ENDIMAGE)\s*\]\([^)]*\)/g, "$1")
				// the spaces that held the word apart from the text around it are
				// not wanted once a picture has a line to itself
				.replace(/^[ \t]+(PDMAILIMAGE\d+ENDIMAGE)[ \t]*$/gm, "$1")
		)
		.replace(/\u00A0/g, " ")
		.replace(/[ \t]+$/gm, "")
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

/** The file types worth their own color, keyed by extension. The colors are
 *  the ones people already read without thinking: Word blue, Excel green,
 *  PowerPoint orange, PDF red. Everything else groups by what it is, so a
 *  bar of attachments is scannable by color before a single name is read. */
// prettier-ignore
const ATT_TYPES: { exts: string[]; color: string }[] = [
	{ exts: ["pdf"], color: "#c0392b" },
	{ exts: ["doc", "docx", "docm", "dot", "dotx", "odt", "rtf", "pages"], color: "#2b579a" },
	{ exts: ["xls", "xlsx", "xlsm", "xlsb", "csv", "tsv", "ods", "numbers"], color: "#217346" },
	{ exts: ["ppt", "pptx", "pptm", "pot", "potx", "odp", "key"], color: "#c43e1c" },
	{ exts: ["png", "jpg", "jpeg", "gif", "webp", "bmp", "heic", "heif", "tif", "tiff", "svg", "avif"], color: "#8250df" },
	{ exts: ["zip", "rar", "7z", "tar", "gz", "bz2", "xz"], color: "#b7791f" },
	{ exts: ["mp4", "mov", "avi", "mkv", "webm", "wmv", "m4v"], color: "#bf3989" },
	{ exts: ["mp3", "wav", "m4a", "aac", "flac", "ogg", "opus", "wma"], color: "#0f766e" },
	{ exts: ["msg", "eml", "oft"], color: "#0f6cbd" },
	{ exts: ["ics", "vcs", "vcf"], color: "#7048c4" },
	{ exts: ["htm", "html", "xml", "json", "yml", "yaml", "js", "ts", "py", "cs", "java", "sql", "sh"], color: "#0e7490" },
	{ exts: ["txt", "md", "log", "ini", "cfg"], color: "#57606a" },
];

/** The MIME families worth recognizing when a name carries no extension.
 *  Graph always sends a contentType, so this is the fallback that keeps a
 *  bare "image001" from landing in the gray bucket. */
// prettier-ignore
const ATT_MIME: { test: RegExp; label: string; color: string }[] = [
	{ test: /^image\//i, label: "IMG", color: "#8250df" },
	{ test: /^video\//i, label: "VID", color: "#bf3989" },
	{ test: /^audio\//i, label: "AUD", color: "#0f766e" },
	{ test: /pdf/i, label: "PDF", color: "#c0392b" },
	{ test: /(zip|compressed|tar)/i, label: "ZIP", color: "#b7791f" },
	{ test: /(rfc822|outlook)/i, label: "MAIL", color: "#0f6cbd" },
	{ test: /calendar/i, label: "ICS", color: "#7048c4" },
	{ test: /^text\//i, label: "TXT", color: "#57606a" },
];

/** An attachment's badge: the short label to print on it and the color to
 *  print it in. The label is the extension, which is what the eye is looking
 *  for anyway, capped at four characters so a long one cannot stretch the
 *  chip. A type nothing recognizes still gets its extension on a neutral
 *  gray, which reads better than a generic paperclip. */
export function attachmentBadge(name: string, contentType = ""): { label: string; color: string } {
	const ext = (name.match(/\.([A-Za-z0-9]{1,8})$/)?.[1] ?? "").toLowerCase();
	if (ext) {
		const hit = ATT_TYPES.find((t) => t.exts.includes(ext));
		return { label: ext.slice(0, 4).toUpperCase(), color: hit?.color ?? "#6b7280" };
	}
	const mime = ATT_MIME.find((t) => t.test.test(contentType));
	if (mime) return { label: mime.label, color: mime.color };
	return { label: "FILE", color: "#6b7280" };
}

// prettier-ignore
const MIME_BY_EXT: Record<string, string> = {
	pdf: "application/pdf", txt: "text/plain", md: "text/markdown", csv: "text/csv", tsv: "text/tab-separated-values",
	html: "text/html", htm: "text/html", xml: "application/xml", json: "application/json",
	png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", svg: "image/svg+xml",
	bmp: "image/bmp", avif: "image/avif", heic: "image/heic", tif: "image/tiff", tiff: "image/tiff",
	doc: "application/msword", docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	xls: "application/vnd.ms-excel", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	ppt: "application/vnd.ms-powerpoint", pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
	zip: "application/zip", gz: "application/gzip", tar: "application/x-tar", "7z": "application/x-7z-compressed",
	mp3: "audio/mpeg", wav: "audio/wav", m4a: "audio/mp4", ogg: "audio/ogg", flac: "audio/flac",
	mp4: "video/mp4", webm: "video/webm", mov: "video/quicktime",
	ics: "text/calendar", eml: "message/rfc822",
};

/** A file extension's content type, for outgoing attachments. Graph accepts
 *  a generic type happily, but the right one is what decides whether the
 *  recipient's client offers to preview the file or only to download it. */
export function mimeForExtension(ext: string): string {
	return MIME_BY_EXT[(ext ?? "").toLowerCase().replace(/^\./, "")] ?? "application/octet-stream";
}

/** An attachment's size the way a mail client writes it: whole KB up to a
 *  megabyte, one decimal past it, and never "0 KB" for a file that exists. */
export function fmtAttachmentSize(bytes: number): string {
	if (bytes >= 1073741824) return `${(bytes / 1073741824).toFixed(1)} GB`;
	if (bytes >= 1048576) return `${(bytes / 1048576).toFixed(1)} MB`;
	return `${Math.max(1, Math.round(bytes / 1024))} KB`;
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
	// a description written in the rich editor arrives as markup and has to
	// be declared as such, or Graph shows the tags to everyone invited;
	// anything without a tag in it still goes as plain text
	if (d.description != null) body.body = { contentType: /<[a-z][\s\S]*>/i.test(d.description) ? "html" : "text", content: d.description };
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
	headers?: Record<string, string>;
	body?: unknown;
}

/** The body of a $batch that rewrites several messages' categories.
 *
 *  A write batch needs what a read batch does not: a Content-Type on each
 *  sub-request, since Graph parses each one as its own request and a PATCH
 *  with no content type is rejected. Sub-requests are keyed by index, as the
 *  read batch does, because Graph ids are far too long for the field. */
export function buildCategoryPatchBatch(items: readonly { id: string; categories: string[] }[]): { requests: BatchRequest[] } {
	return {
		requests: items.map((it, i) => ({
			id: String(i),
			method: "PATCH",
			url: `/me/messages/${encodeURIComponent(it.id)}`,
			headers: { "Content-Type": "application/json" },
			body: { categories: it.categories },
		})),
	};
}

/** The body of a $batch that marks several messages read or unread. Same
 *  shape as the category batch and for the same reason: Graph has no bulk
 *  read API, so a folder of four hundred unread is four hundred writes, and
 *  twenty at a time is the difference between a second and a minute. */
export function buildReadPatchBatch(ids: readonly string[], read: boolean): { requests: BatchRequest[] } {
	return {
		requests: ids.map((id, i) => ({
			id: String(i),
			method: "PATCH",
			url: `/me/messages/${encodeURIComponent(id)}`,
			headers: { "Content-Type": "application/json" },
			body: { isRead: read },
		})),
	};
}

/** Which ids in a write batch came back happy, by index. */
export function parseWriteBatch(reply: unknown, ids: readonly string[]): { ok: string[]; failed: string[] } {
	const responses = (reply as { responses?: { id?: string; status?: number }[] } | null)?.responses;
	if (!Array.isArray(responses)) return { ok: [], failed: [...ids] };
	const ok: string[] = [];
	const seen = new Set<string>();
	for (const r of responses) {
		const idx = Number(r?.id);
		const id = ids[idx];
		if (id === undefined) continue;
		seen.add(id);
		if ((r?.status ?? 500) < 400) ok.push(id);
	}
	return { ok, failed: ids.filter((id) => !ok.includes(id)) };
}

/** Replace one category name with another in a list, leaving the rest and
 *  their order alone. Used to retag a message during a category replace. */
export function replaceCategory(current: string[] | undefined, from: string, to: string): string[] {
	const seen = new Set<string>();
	const out: string[] = [];
	for (const c of current ?? []) {
		const next = c.toLowerCase() === from.toLowerCase() ? to : c;
		// a message already carrying both names must not end up with the new
		// one twice, which Outlook would show as a duplicate label
		if (seen.has(next.toLowerCase())) continue;
		seen.add(next.toLowerCase());
		out.push(next);
	}
	return out;
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

/* ---------------- keyed list reconciliation ---------------- */

/** The slice of a DOM node this reconciler needs. Structural and generic in
 *  its own sibling type, so the browser's Element satisfies it exactly and a
 *  test can stand in a plain object without a cast. */
export interface ReconcileNode<N> {
	readonly isConnected: boolean;
	readonly nextElementSibling: N | null;
	remove(): void;
}

/** The slice of a DOM container this reconciler needs. */
export interface ReconcileParent<N> {
	readonly firstElementChild: N | null;
	insertBefore(node: N, before: N | null): unknown;
	empty(): void;
}

/** One thing that can appear in the column, and how to build it if it must. */
export interface ReconcileItem<N> {
	/** Stable across renders: the message id, the section name. Rows keyed by
	 *  identity are what let mail arriving at the top insert one node instead
	 *  of shifting every later one's description down and remaking the lot. */
	key: string;
	/** Everything the node would draw. Equal means leave it alone. */
	sig: string;
	/** Builds the node, APPENDED to the parent; the reconciler moves it. */
	make: () => N;
}

export function reconcileChildren<N extends ReconcileNode<N>>(
	parent: ReconcileParent<N>,
	items: readonly ReconcileItem<N>[],
	prev: Map<string, { sig: string; el: N }>
): Map<string, { sig: string; el: N }> {
	const next = new Map<string, { sig: string; el: N }>();
	// nothing on screen survives (changing folder, say), and clearing the whole
	// column at once beats unpicking it a node at a time
	let anyKnown = false;
	for (const item of items) {
		if (prev.has(item.key)) {
			anyKnown = true;
			break;
		}
	}
	if (!anyKnown) parent.empty();
	let cursor: N | null = parent.firstElementChild;
	for (const item of items) {
		// two items can only share a key if the same message is listed twice,
		// which grouping should never produce; disambiguate rather than let the
		// second one quietly replace the first
		let key = item.key;
		for (let n = 2; next.has(key); n++) key = `${item.key}#${n}`;
		const old = prev.get(key);
		let el: N;
		if (old && old.sig === item.sig && old.el.isConnected) el = old.el;
		else {
			// step the cursor off a node about to be removed: a detached cursor
			// has no next sibling left to walk to
			if (old?.el && old.el === cursor) cursor = cursor.nextElementSibling;
			old?.el.remove();
			el = item.make();
		}
		if (el === cursor) cursor = cursor.nextElementSibling;
		else parent.insertBefore(el, cursor);
		next.set(key, { sig: item.sig, el });
	}
	// whatever the new list has no place for
	while (cursor) {
		const after = cursor.nextElementSibling;
		cursor.remove();
		cursor = after;
	}
	return next;
}
