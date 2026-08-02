/* iCalendar parsing and recurrence expansion over ical.js. Pure: no Obsidian
 * imports, unit-tested with Node. Both the ICS-feed source and the CalDAV
 * source funnel their payloads through here, so every provider's RRULE,
 * EXDATE, and moved-instance quirks are handled once. */
import ICAL from "ical.js";
import { PCEvent, findJoinUrl } from "./core";

export interface IcsSourceInfo {
	sourceId: string;
	calendarName?: string;
	color?: string;
}

/** Safety valve: a series is never expanded past this many steps, so a
 *  malformed every-minute rule cannot hang a refresh. */
const MAX_ITERATIONS = 10000;

/** All concrete event instances in one iCalendar payload that intersect
 *  [winStartMs, winEndMs). Recurring series expand client-side; EXDATEs and
 *  cancelled or moved overrides are honoured. Unparseable payloads throw so
 *  the source can surface the error; a single busted VEVENT is skipped. */
export function parseIcsEvents(text: string, winStartMs: number, winEndMs: number, src: IcsSourceInfo): PCEvent[] {
	const comp = new ICAL.Component(ICAL.parse(text));

	// Register the payload's own timezone definitions. Outlook feeds reference
	// Windows zone names ("Tokyo Standard Time"); their VTIMEZONE blocks define
	// the offsets, so registration is what makes those times resolvable.
	ICAL.TimezoneService.reset();
	for (const vtz of comp.getAllSubcomponents("vtimezone")) {
		try {
			ICAL.TimezoneService.register(new ICAL.Timezone(vtz));
		} catch {
			/* an unreadable zone falls back to floating time below */
		}
	}

	// Group a series' master with its overrides (RECURRENCE-ID VEVENTs).
	const mains = new Map<string, ICAL.Event>();
	const exceptions: ICAL.Event[] = [];
	for (const v of comp.getAllSubcomponents("vevent")) {
		let ev: ICAL.Event;
		try {
			ev = new ICAL.Event(v);
		} catch {
			continue;
		}
		if (ev.isRecurrenceException()) exceptions.push(ev);
		else mains.set(ev.uid || `#${mains.size}`, ev);
	}
	for (const ex of exceptions) {
		const main = mains.get(ex.uid || "");
		if (main) {
			try {
				main.relateException(ex);
			} catch {
				/* an unrelatable override renders as its own single event */
				mains.set(`${ex.uid}#${String(ex.recurrenceId)}`, ex);
			}
		} else {
			// CalDAV REPORTs can return an override whose master fell outside the
			// query; the moved instance still deserves to render.
			mains.set(`${ex.uid}#${String(ex.recurrenceId)}`, ex);
		}
	}

	const out: PCEvent[] = [];
	for (const ev of mains.values()) {
		try {
			expandEvent(ev, winStartMs, winEndMs, src, out);
		} catch {
			/* one busted series must not take the whole feed down */
		}
	}
	return out;
}

function expandEvent(ev: ICAL.Event, winStartMs: number, winEndMs: number, src: IcsSourceInfo, out: PCEvent[]): void {
	if (!ev.isRecurring()) {
		const s = timeToMs(ev.startDate);
		const e = endMsOf(ev.startDate, ev.endDate, s);
		if (s < winEndMs && e > winStartMs && !isCancelled(ev)) {
			out.push(instance(ev, ev, s, e, !!ev.startDate.isDate, false, src));
		}
		return;
	}

	const it = ev.iterator();
	let next: ICAL.Time | null;
	let guard = 0;
	while ((next = it.next()) && guard++ < MAX_ITERATIONS) {
		// The recurrence id walks in order, so once IT passes the window we are
		// done. Occurrence times decide inclusion (an override may have moved).
		if (timeToMs(next) >= winEndMs) break;
		let det: ReturnType<ICAL.Event["getOccurrenceDetails"]>;
		try {
			det = ev.getOccurrenceDetails(next);
		} catch {
			continue;
		}
		const item = det.item; // the override event when this occurrence has one
		if (isCancelled(item)) continue;
		const s = timeToMs(det.startDate);
		const e = endMsOf(det.startDate, det.endDate, s);
		if (s >= winEndMs || e <= winStartMs) continue;
		out.push(instance(ev, item, s, e, !!det.startDate.isDate, true, src));
	}
}

/** Local epoch ms of an ICAL.Time. Date-only values build from their own
 *  Y/M/D as LOCAL midnight, bypassing whatever zone ical.js attached, so an
 *  all-day event can never shift a day on a negative-offset machine. */
function timeToMs(t: ICAL.Time): number {
	if (t.isDate) return new Date(t.year, t.month - 1, t.day).getTime();
	return t.toJSDate().getTime();
}

function endMsOf(start: ICAL.Time, end: ICAL.Time | null | undefined, startMs: number): number {
	const e = end ? timeToMs(end) : startMs;
	if (e > startMs) return e;
	return start.isDate ? startMs + 86400000 : startMs; // DTEND-less dates are one day
}

function isCancelled(ev: ICAL.Event): boolean {
	const status = ev.component.getFirstPropertyValue("status");
	return typeof status === "string" && status.toUpperCase() === "CANCELLED";
}

function instance(main: ICAL.Event, item: ICAL.Event, startMs: number, endMs: number, isDate: boolean, recurring: boolean, src: IcsSourceInfo): PCEvent {
	const location = item.location || undefined;
	const description = item.description || undefined;
	const urlProp = item.component.getFirstPropertyValue("url");
	const url = urlProp == null ? undefined : String(urlProp);
	return {
		id: `${main.uid || "?"}:${startMs}`,
		sourceId: src.sourceId,
		calendarName: src.calendarName,
		color: src.color,
		title: (item.summary || main.summary || "").trim() || "(no title)",
		allDay: isDate,
		startMs,
		endMs,
		location,
		organizer: personName(item.component.getFirstProperty("organizer")),
		attendees: attendeeNames(item),
		description,
		url,
		joinUrl: findJoinUrl(location, description, url) || undefined,
		recurring,
		tentative: String(item.component.getFirstPropertyValue("status") ?? "").toUpperCase() === "TENTATIVE",
	};
}

/** CN parameter when present, otherwise the mailto address without its scheme. */
function personName(prop: ICAL.Property | null): string | undefined {
	if (!prop) return undefined;
	const cn = prop.getParameter("cn");
	if (typeof cn === "string" && cn.trim()) return cn.trim();
	const v = String(prop.getFirstValue() ?? "").replace(/^mailto:/i, "");
	return v || undefined;
}

function attendeeNames(item: ICAL.Event): string[] | undefined {
	const out: string[] = [];
	for (const p of item.component.getAllProperties("attendee")) {
		const kind = p.getParameter("cutype");
		if (typeof kind === "string" && /^(resource|room)$/i.test(kind)) continue;
		const name = personName(p);
		if (name) out.push(name);
	}
	return out.length ? out : undefined;
}
