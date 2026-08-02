/* Unit tests over the pure modules (core.ts, ics.ts, caldavxml.ts).
 * Run via: npm test. No Obsidian, no network, no framework. */
import {
	PCEvent,
	PCMail,
	padFetchWindow,
	addDays,
	b64utf8,
	clipToDay,
	dayDiff,
	dayOfWeek,
	daySpanKeys,
	decodeJwtPayload,
	dedupePeople,
	dragTimes,
	dueReminders,
	eventDaySpan,
	eventsOnDay,
	findJoinUrl,
	meetingProvider,
	fmtClock,
	fmtDayHeading,
	fmtEventRange,
	fmtFreeSlots,
	fmtZoneClock,
	freeSlotsForDay,
	fmtMailTime,
	folderSubtreeIds,
	googleEventBody,
	googleRecurrence,
	graphFolderToPC,
	graphMailToPC,
	mergeDeltaMessages,
	orderFolderTree,
	subjectToEventTitle,
	googleTimesBody,
	googleToPC,
	graphEventBody,
	graphRecurrence,
	graphSetupHint,
	graphTimesBody,
	graphToPC,
	groupByDay,
	isoWeekNum,
	isSpanEvent,
	keyOfMs,
	mergeForSave,
	minutesOfMs,
	monthGrid,
	msOfKey,
	dateKeyOf,
	packColumns,
	packLanes,
	paletteColor,
	parseAgendaBlock,
	parseAvailabilityView,
	slotConflict,
	avatarColor,
	avatarInitials,
	PALETTE,
	pickGeoHit,
	searchFolderQuery,
	splitPlaceQuery,
	stripMeetingBoilerplate,
	weatherGlyph,
	parseAttendeeInput,
	parseWallClock,
	periodLabel,
	renderNoteName,
	sanitizeName,
	snapMin,
	sortEvents,
	spansForRow,
	stepAnchor,
	stripHtml,
	timedOnDay,
	vaultDateSpan,
	fmtVaultDate,
	viewWindow,
	wallOfMs,
	weekDays,
	workWeekDays,
} from "./core";

let failures = 0;
function eq(a: unknown, b: unknown, msg: string) {
	const sa = JSON.stringify(a);
	const sb = JSON.stringify(b);
	if (sa === sb) console.log("  ok -", msg);
	else {
		failures++;
		console.error("  FAIL -", msg, "\n    got:     ", sa, "\n    expected:", sb);
	}
}

/** A timed event built from LOCAL wall-clock parts, so every assertion below
 *  is timezone-independent. */
function evAt(key: string, startMin: number, durMin: number, extra: Partial<PCEvent> = {}): PCEvent {
	const startMs = msOfKey(key) + startMin * 60000;
	return {
		id: `${key}:${startMin}`,
		sourceId: "t",
		title: "T",
		allDay: false,
		startMs,
		endMs: startMs + durMin * 60000,
		...extra,
	};
}

function allDayEv(fromKey: string, days: number, extra: Partial<PCEvent> = {}): PCEvent {
	return {
		id: `ad:${fromKey}`,
		sourceId: "t",
		title: "AD",
		allDay: true,
		startMs: msOfKey(fromKey),
		endMs: msOfKey(addDays(fromKey, days)),
		...extra,
	};
}

// --- day-key algebra ---
eq(addDays("2026-07-17", 1), "2026-07-18", "addDays steps forward");
eq(addDays("2026-01-01", -1), "2025-12-31", "addDays crosses a year");
eq(addDays("2024-02-28", 1), "2024-02-29", "addDays knows leap years");
eq(dayDiff("2026-07-17", "2026-07-20"), 3, "dayDiff positive when b later");
eq(dayOfWeek("2026-07-17"), 5, "2026-07-17 is a Friday");
eq(weekDays("2026-07-17", true)[0], "2026-07-13", "monday-start week begins Monday");
eq(weekDays("2026-07-17", false)[0], "2026-07-12", "sunday-start week begins Sunday");
eq(weekDays("2026-07-13", true)[0], "2026-07-13", "a Monday starts its own monday week");

// --- monthGrid ---
{
	const july = monthGrid(2026, 6, true);
	eq(july.length, 42, "month grid is always 42 cells");
	eq(july[0].key, "2026-06-29", "monday-start July 2026 leads with prior Monday");
	eq(july[2].inMonth, true, "July 1 cell is in-month");
	eq(july[0].inMonth, false, "lead cells are out-of-month");
	const sun = monthGrid(2026, 6, false);
	eq(sun[0].key, "2026-06-28", "sunday-start July 2026 leads with prior Sunday");
}

// --- local bridges ---
eq(keyOfMs(msOfKey("2026-07-17")), "2026-07-17", "msOfKey and keyOfMs round-trip");
eq(minutesOfMs(msOfKey("2026-07-17") + 9.5 * 3600000), 570, "minutesOfMs reads local wall clock");

// --- event geometry ---
eq(eventDaySpan(evAt("2026-07-17", 540, 60)), { startKey: "2026-07-17", endKey: "2026-07-17" }, "one-hour event spans one day");
eq(eventDaySpan(evAt("2026-07-17", 1320, 180)), { startKey: "2026-07-17", endKey: "2026-07-18" }, "10pm-1am event spans two days");
eq(eventDaySpan(evAt("2026-07-17", 1380, 60)), { startKey: "2026-07-17", endKey: "2026-07-17" }, "event ending exactly at midnight stays on its day");
eq(eventDaySpan(allDayEv("2026-07-17", 1)), { startKey: "2026-07-17", endKey: "2026-07-17" }, "single all-day spans one day (exclusive end)");
eq(eventDaySpan(allDayEv("2026-07-17", 3)), { startKey: "2026-07-17", endKey: "2026-07-19" }, "3-day all-day ends on its last day");
eq(isSpanEvent(evAt("2026-07-17", 540, 60)), false, "a timed one-day event is a block");
eq(isSpanEvent(evAt("2026-07-17", 1320, 180)), true, "a midnight-crossing event is a banner");
eq(isSpanEvent(allDayEv("2026-07-17", 1)), true, "an all-day event is a banner");

eq(clipToDay(evAt("2026-07-17", 540, 60), "2026-07-17"), { startMin: 540, endMin: 600 }, "clip inside the day is identity");
eq(clipToDay(evAt("2026-07-17", 1320, 180), "2026-07-17"), { startMin: 1320, endMin: 1440 }, "clip cuts at midnight going out");
eq(clipToDay(evAt("2026-07-17", 1320, 180), "2026-07-18"), { startMin: 0, endMin: 60 }, "clip starts at midnight coming in");
eq(clipToDay(evAt("2026-07-17", 540, 60), "2026-07-18"), null, "clip is null on untouched days");
eq(clipToDay(evAt("2026-07-17", 600, 0), "2026-07-17"), { startMin: 600, endMin: 615 }, "zero-length events get a visible sliver");

{
	const evs = [evAt("2026-07-17", 540, 60), evAt("2026-07-18", 540, 60), allDayEv("2026-07-17", 2)];
	eq(eventsOnDay(evs, "2026-07-17").length, 2, "eventsOnDay finds timed and all-day");
	eq(eventsOnDay(evs, "2026-07-18").length, 2, "multi-day all-day appears on its second day");
	eq(eventsOnDay(evs, "2026-07-19").length, 0, "exclusive all-day end does not leak a day");
}

// --- sorting ---
{
	const a = evAt("2026-07-17", 600, 30, { title: "b" });
	const b = evAt("2026-07-17", 540, 30, { title: "a" });
	const banner = allDayEv("2026-07-17", 1);
	eq(
		[a, b, banner].sort(sortEvents).map((e) => e.title),
		["AD", "a", "b"],
		"sort puts banners first, then start time"
	);
}

// --- packColumns ---
{
	const disjoint = packColumns([
		{ startMin: 0, endMin: 60 },
		{ startMin: 60, endMin: 120 },
	]);
	eq(disjoint, [
		{ col: 0, cols: 1 },
		{ col: 0, cols: 1 },
	], "back-to-back events do not share a cluster");
	const pair = packColumns([
		{ startMin: 0, endMin: 60 },
		{ startMin: 30, endMin: 90 },
	]);
	eq(pair, [
		{ col: 0, cols: 2 },
		{ col: 1, cols: 2 },
	], "overlapping pair splits into two columns");
	const chain = packColumns([
		{ startMin: 0, endMin: 60 },
		{ startMin: 30, endMin: 90 },
		{ startMin: 75, endMin: 120 },
	]);
	eq(chain, [
		{ col: 0, cols: 2 },
		{ col: 1, cols: 2 },
		{ col: 0, cols: 2 },
	], "a chain reuses freed columns and shares the cluster width");
	const twoClusters = packColumns([
		{ startMin: 0, endMin: 30 },
		{ startMin: 10, endMin: 40 },
		{ startMin: 100, endMin: 130 },
	]);
	eq(twoClusters[2], { col: 0, cols: 1 }, "a later lone event starts a fresh cluster");
}

// --- packLanes / spansForRow ---
eq(packLanes([
	{ start: 0, end: 2 },
	{ start: 1, end: 3 },
	{ start: 3, end: 5 },
]), [0, 1, 0], "lanes: overlap stacks, a freed lane is reused");
{
	const row = weekDays("2026-07-17", true); // Jul 13..19
	const spans = spansForRow(
		[allDayEv("2026-07-12", 3), allDayEv("2026-07-15", 2), evAt("2026-07-17", 540, 60)],
		row
	);
	eq(spans.length, 2, "spansForRow keeps banners only");
	eq(spans[0].startIdx, 0, "a span from last week clips to column 0");
	eq(spans[0].startsHere, false, "and knows it continued in");
	eq(spans[0].endIdx, 1, "3-day span from Jul 12 ends in column 1 (Jul 14)");
	eq(spans[1].startIdx, 2, "Jul 15 span starts in column 2");
	eq([spans[0].lane, spans[1].lane], [0, 0], "non-overlapping spans share lane 0");
}

// --- timedOnDay / groupByDay ---
{
	const evs = [evAt("2026-07-17", 540, 60), evAt("2026-07-17", 1320, 180), allDayEv("2026-07-17", 1)];
	eq(timedOnDay(evs, "2026-07-17").length, 1, "timedOnDay excludes banners and midnight-crossers");
	const groups = groupByDay(evs, "2026-07-16", "2026-07-19");
	eq(groups.map((g) => g.key), ["2026-07-17", "2026-07-18"], "agenda groups skip empty days");
	eq(groups[0].events.length, 3, "agenda lists every event touching the day");
}

// --- windows and navigation ---
eq(viewWindow("month", "2026-07-17", true), { fromKey: "2026-06-29", toKey: "2026-08-09" }, "month window is the 42-cell grid");
eq(viewWindow("week", "2026-07-17", true), { fromKey: "2026-07-13", toKey: "2026-07-19" }, "week window");
eq(viewWindow("day", "2026-07-17", true), { fromKey: "2026-07-17", toKey: "2026-07-17" }, "day window");
eq(viewWindow("agenda", "2026-07-17", true, 30), { fromKey: "2026-07-17", toKey: "2026-08-15" }, "agenda window is N days");
eq(stepAnchor("month", "2026-01-31", 1), "2026-02-01", "month step pins to the 1st so short months are not skipped");
eq(stepAnchor("month", "2026-01-15", -1), "2025-12-01", "month step back crosses the year");
eq(stepAnchor("week", "2026-07-17", 1), "2026-07-24", "week step is 7 days");
eq(stepAnchor("day", "2026-07-17", -1), "2026-07-16", "day step is 1 day");
eq(workWeekDays("2026-07-17"), ["2026-07-13", "2026-07-14", "2026-07-15", "2026-07-16", "2026-07-17"], "work week is monday through friday");
eq(workWeekDays("2026-07-19")[0], "2026-07-13", "a sunday belongs to the work week that just ended");
eq(viewWindow("workweek", "2026-07-17", false), { fromKey: "2026-07-13", toKey: "2026-07-17" }, "work week window ignores the week-start setting");
eq(stepAnchor("workweek", "2026-07-17", 1), "2026-07-24", "work week steps by a calendar week");
eq(periodLabel("workweek", "2026-07-17", true), "Jul 13 - 17, 2026", "work week label");

// --- labels ---
eq(fmtClock(570, false), "9:30 AM", "12h clock");
eq(fmtClock(570, true), "09:30", "24h clock");
eq(fmtClock(540, false, true), "9 AM", "short 12h drops :00");
// Agenda rows ask for the long form: a time column reads as a column when every
// row is the same shape, so 7:00 AM lines up under 12:30 AM instead of "7 AM".
eq(fmtClock(540, false), "9:00 AM", "the default keeps :00, for lists that stack times in a column");
eq(fmtClock(0, false), "12:00 AM", "midnight is 12 AM");
eq(fmtClock(720, false), "12:00 PM", "noon is 12 PM");
eq(periodLabel("month", "2026-07-01", true), "July 2026", "month label");
eq(periodLabel("week", "2026-07-17", true), "Jul 13 - 19, 2026", "same-month week label");
eq(periodLabel("week", "2026-07-01", true), "Jun 29 - Jul 5, 2026", "cross-month week label");
eq(periodLabel("week", "2025-12-31", true), "Dec 29, 2025 - Jan 4, 2026", "cross-year week label");
eq(periodLabel("day", "2026-07-17", true), "Friday, Jul 17, 2026", "day label");
eq(fmtDayHeading("2026-07-17"), "Friday, Jul 17", "agenda day heading");
eq(fmtEventRange(evAt("2026-07-17", 540, 30), false), "9:00 AM - 9:30 AM", "timed range label");
eq(fmtEventRange(allDayEv("2026-07-17", 1), false), "All day", "single all-day label");
eq(fmtEventRange(allDayEv("2026-07-17", 3), false), "Jul 17 - Jul 19", "multi-day all-day label");
eq(fmtEventRange(evAt("2026-07-17", 1320, 180), false), "Jul 17, 10:00 PM - Jul 18, 1:00 AM", "midnight-crossing label names both days");

// --- note names ---
eq(sanitizeName('Q3: sync / review <plan> #4 [x]'), "Q3 sync review plan 4 x", "sanitize strips path and link syntax");
eq(sanitizeName("   "), "Untitled", "sanitize never returns empty");
{
	const ev = evAt("2026-07-17", 570, 30, { title: "Team sync: Q3", calendarName: "Work" });
	eq(renderNoteName("{{date}} {{title}}", ev, false), "2026-07-17 Team sync Q3", "note name template");
	eq(renderNoteName("{{date}} {{time}} {{title}} ({{calendar}})", ev, false), "2026-07-17 09.30 Team sync Q3 (Work)", "time token is filename-safe");
	eq(renderNoteName("{{time}}{{title}}", allDayEv("2026-07-17", 1, { title: "Trip" }), false), "Trip", "all-day time token is empty");
}

// --- text utilities ---
eq(
	stripHtml("<div>Agenda:<br>1. Hello &amp; welcome</div><p>Bring &quot;notes&quot;</p>"),
	'Agenda:\n1. Hello & welcome\nBring "notes"',
	"stripHtml makes readable text"
);
eq(findJoinUrl("join here https://teams.microsoft.com/l/meetup-join/abc?x=1 now"), "https://teams.microsoft.com/l/meetup-join/abc?x=1", "finds a Teams link");
eq(findJoinUrl("call at https://us02web.zoom.us/j/123456"), "https://us02web.zoom.us/j/123456", "finds a Zoom link on a subdomain");
eq(findJoinUrl("nothing to join"), null, "no link is null");
eq(meetingProvider("https://us02web.zoom.us/j/8786?pwd=8LtS653eVLSEW4"), "Zoom", "a zoom subdomain is named, not printed");
eq(meetingProvider("https://teams.microsoft.com/l/meetup-join/abc"), "Teams", "Teams is named");
eq(meetingProvider("https://meet.google.com/abc-defg-hij"), "Google Meet", "Meet is named");
eq(meetingProvider("https://acme.webex.com/j/1"), "Webex", "a hosted Webex is named");
eq(meetingProvider("https://www.chime.aws/1234"), "chime.aws", "an unknown host falls back to its domain, never the full link");
eq(meetingProvider(""), "", "no link, nothing to name");
eq(meetingProvider("not a url"), "", "an unparseable link names nothing rather than guessing");
eq(b64utf8("user:pass"), "dXNlcjpwYXNz", "basic-auth base64");
eq(b64utf8("ü:π"), "w7w6z4A=", "base64 survives non-ascii");
eq(paletteColor(11), paletteColor(1), "palette wraps");

// --- graphToPC ---
{
	const raw = {
		id: "AAA",
		subject: "Standup",
		isAllDay: false,
		showAs: "tentative",
		type: "occurrence",
		start: { dateTime: "2026-07-17T09:00:00.0000000", timeZone: "X" },
		end: { dateTime: "2026-07-17T09:30:00.0000000", timeZone: "X" },
		location: { displayName: "Room 4" },
		organizer: { emailAddress: { name: "Ana", address: "ana@x.com" } },
		attendees: [
			{ type: "required", emailAddress: { name: "Bob", address: "bob@x.com" } },
			{ type: "resource", emailAddress: { name: "Room 4", address: "room4@x.com" } },
		],
		onlineMeeting: { joinUrl: "https://teams.microsoft.com/l/m/1" },
		webLink: "https://outlook.office365.com/owa/?id=1",
		bodyPreview: "notes",
		responseStatus: { response: "accepted" },
	};
	const pc = graphToPC(raw, { sourceId: "m365", calendarName: "Calendar", color: "#123456" });
	eq(pc?.title, "Standup", "graph event maps title");
	eq(pc?.startMs, msOfKey("2026-07-17") + 540 * 60000, "graph wall time reads as local");
	eq(pc?.attendees, ["Bob"], "resource attendees are dropped");
	eq(pc?.tentative, true, "showAs tentative maps");
	eq(pc?.recurring, true, "occurrence type marks recurring");
	eq(graphToPC({ ...raw, seriesMasterId: "SM1" }, { sourceId: "m365" })?.seriesId, "SM1", "the graph series master id rides along");
	eq(pc?.joinUrl, "https://teams.microsoft.com/l/m/1", "online meeting join url wins");
	eq(graphToPC({ ...raw, isCancelled: true }, { sourceId: "m365" }), null, "cancelled maps to null");
	const allDay = graphToPC(
		{ id: "B", subject: "PTO", isAllDay: true, start: { dateTime: "2026-07-17T00:00:00.0000000" }, end: { dateTime: "2026-07-18T00:00:00.0000000" } },
		{ sourceId: "m365" }
	);
	eq(allDay && eventDaySpan(allDay), { startKey: "2026-07-17", endKey: "2026-07-17" }, "graph all-day lands on one local day");
	const declined = graphToPC({ ...raw, responseStatus: { response: "declined" } }, { sourceId: "m365" });
	eq(declined?.declined, true, "declined response maps");
}
eq(parseWallClock("garbage"), null, "unparseable wall clock is null");
eq(graphSetupHint("AADSTS7000218: public client flows off")?.includes("Allow public client flows"), true, "setup hint names the public-client fix");
eq(graphSetupHint("AADSTS500200: personal account not supported")?.includes("own app registration"), true, "personal-account refusals point at the wizard's personal path");
eq(graphSetupHint("ErrorAccessDenied: Access is denied. Check credentials and try again.")?.includes("Reconnect"), true, "a permission gap points at the reconnect button");
eq(graphSetupHint("something else entirely"), null, "no known code, no hint");

// --- search folder types ---
eq(searchFolderQuery("unread", "", "me@x.com"), "isread:false", "unread compiles");
eq(searchFolderQuery("unreadOrFlagged", "", "me@x.com"), "isread:false OR isflagged:true", "the either type ORs both");
eq(searchFolderQuery("toMe", "", "me@x.com"), "to:me@x.com", "to-me carries the mailbox address");
eq(searchFolderQuery("fromPeople", "ana@x.com, bob@y.com", "me@x.com"), "from:ana@x.com OR from:bob@y.com", "people OR together");
eq(searchFolderQuery("fromToPeople", "ana@x.com", "me@x.com"), "from:ana@x.com OR to:ana@x.com", "from-and-to covers both directions");
eq(searchFolderQuery("category", "Red category", "me@x.com"), 'category:"Red category"', "categories quote their name");
eq(searchFolderQuery("custom", "subject:invoice", "me@x.com"), "subject:invoice", "custom passes the raw query");

// --- mail avatars ---
eq(avatarInitials("Uber Receipts"), "UR", "two words give two letters");
eq(avatarInitials("Cursor"), "C", "one word gives one letter");
eq(avatarInitials("alex.kim@example.com"), "AK", "an address reads its local-part segments");
eq(avatarInitials(""), "?", "empty falls back to a question mark");
eq(avatarColor("Uber Receipts"), avatarColor("Uber Receipts"), "the avatar color is stable");
eq(PALETTE.includes(avatarColor("Anyone At All")), true, "avatar colors come from the palette");

// --- place lookup ---
{
	eq(splitPlaceQuery("Flower Mound, TX"), { name: "Flower Mound", region: "TX" }, "the comma form splits into city and region");
	eq(splitPlaceQuery("Fort Wayne"), { name: "Fort Wayne", region: "" }, "no comma, no region");
	const hits = [
		{ name: "Chicago", admin1: "Illinois", country_code: "US", latitude: 41.85, longitude: -87.65 },
		{ name: "Chicago", admin1: "Indiana", country_code: "US", latitude: 41, longitude: -87 },
	];
	eq(pickGeoHit("IL", hits)?.latitude, 41.85, "a state abbreviation picks its state's hit");
	eq(pickGeoHit("Indiana", hits)?.latitude, 41, "a full state name matches too");
	eq(pickGeoHit("", hits)?.latitude, 41.85, "no region takes the first hit");
	eq(pickGeoHit("ZZ", hits)?.latitude, 41.85, "an unknown region falls back to the first hit");
	eq(pickGeoHit("TX", []), null, "no hits, no pick");
}

// --- mail flags ---
{
	const m2 = graphMailToPC(
		{ id: "M2", subject: "s", from: { emailAddress: { address: "x@y.com" } }, receivedDateTime: "2026-07-17T12:00:00Z", importance: "high", flag: { flagStatus: "flagged" }, toRecipients: [{ emailAddress: { address: "STEVE@example.com" } }] },
		"a1",
		"Work",
		"steve@example.com"
	);
	eq(m2?.priority, true, "high importance maps to priority");
	eq(m2?.flagged, true, "a follow-up flag maps");
	eq(m2?.toMe, true, "the To line matches the mailbox address case-blind");
	const m3 = graphMailToPC({ id: "M3", subject: "s", receivedDateTime: "2026-07-17T12:00:00Z", toRecipients: [{ emailAddress: { address: "other@y.com" } }] }, "a1", "Work", "steve@example.com");
	eq(m3?.toMe, undefined, "mail not addressed to the mailbox stays unmarked");
}

// --- weather glyphs ---
eq(weatherGlyph(0), "☀️", "clear sky is sun");
eq(weatherGlyph(3), "☁️", "overcast is cloud");
eq(weatherGlyph(61), "🌧️", "rain codes map to rain");
eq(weatherGlyph(96), "⛈️", "thunderstorm codes map to storm");

// --- meeting boilerplate ---
eq(stripMeetingBoilerplate("Agenda first\n\n________________________________________\nMicrosoft Teams meeting\nJoin: https://teams.example"), "Agenda first", "teams boilerplate drops from the separator line");
eq(stripMeetingBoilerplate("Microsoft Teams meeting\nJoin: x\nMeeting ID: 1"), "", "a body that is only boilerplate empties out");
eq(stripMeetingBoilerplate("Join Zoom Meeting\nhttps://zoom.example"), "", "zoom blocks drop too");
eq(stripMeetingBoilerplate("Plain agenda text\nwith two lines"), "Plain agenda text\nwith two lines", "a human body passes through untouched");

// --- multi-day day view ---
eq(viewWindow("day", "2026-07-18", true), { fromKey: "2026-07-18", toKey: "2026-07-18" }, "day view defaults to a single day");
eq(viewWindow("day", "2026-07-18", true, 30, 3), { fromKey: "2026-07-18", toKey: "2026-07-20" }, "day view spans the chosen days");
eq(daySpanKeys("2026-07-30", 3), ["2026-07-30", "2026-07-31", "2026-08-01"], "the day span crosses a month edge");
eq(daySpanKeys("2026-07-18", 99).length, 7, "the span clamps to a week");
eq(stepAnchor("day", "2026-07-18", 1, 30, 3), "2026-07-21", "stepping a 3-day view moves 3 days");
eq(stepAnchor("day", "2026-07-18", -1, 30, 1), "2026-07-17", "a single day steps 1");
eq(periodLabel("day", "2026-07-18", true, 30, 3), "Jul 18 - 20, 2026", "a multi-day label reads as a range");

// --- write-back geometry ---
eq(snapMin(37), 30, "snap rounds to the nearer step");
eq(snapMin(38), 45, "snap rounds up past the midpoint");
eq(snapMin(50, 30), 60, "snap honors a custom step");
{
	const ev = evAt("2026-07-17", 540, 60); // 9:00-10:00
	eq(dragTimes(ev.startMs, ev.endMs, "2026-07-17", 33, "move"), { startMs: msOfKey("2026-07-17") + 570 * 60000, endMs: msOfKey("2026-07-17") + 630 * 60000 }, "move snaps the start and keeps the duration");
	eq(dragTimes(ev.startMs, ev.endMs, "2026-07-18", 0, "move"), { startMs: msOfKey("2026-07-18") + 540 * 60000, endMs: msOfKey("2026-07-18") + 600 * 60000 }, "cross-day move keeps the wall time");
	eq(dragTimes(ev.startMs, ev.endMs, "2026-07-17", -55, "resize"), { startMs: ev.startMs, endMs: ev.startMs + 15 * 60000 }, "resize never shrinks under 15 minutes");
	eq(dragTimes(ev.startMs, ev.endMs, "2026-07-17", 34, "resize"), { startMs: ev.startMs, endMs: ev.startMs + 90 * 60000 }, "resize snaps the new duration");
}
eq(wallOfMs(msOfKey("2026-07-17") + 570 * 60000), "2026-07-17T09:30:00", "wall clock is local and Graph-shaped");
{
	const t = graphTimesBody(msOfKey("2026-07-17") + 540 * 60000, msOfKey("2026-07-17") + 600 * 60000, false, "America/Indiana/Indianapolis");
	eq(t.start, { dateTime: "2026-07-17T09:00:00", timeZone: "America/Indiana/Indianapolis" }, "times body carries the zone");
	const ad = graphTimesBody(msOfKey("2026-07-17"), msOfKey("2026-07-18"), true, "UTC");
	eq((ad.start as { dateTime: string }).dateTime, "2026-07-17T00:00:00", "all-day starts at midnight");
	eq((ad.end as { dateTime: string }).dateTime, "2026-07-18T00:00:00", "all-day end stays the exclusive next midnight");
	const zero = graphTimesBody(msOfKey("2026-07-17"), msOfKey("2026-07-17"), true, "UTC");
	eq((zero.end as { dateTime: string }).dateTime, "2026-07-18T00:00:00", "a zero-length all-day draft becomes one day");
	const body = graphEventBody({ title: "  Standup  ", startMs: msOfKey("2026-07-17") + 540 * 60000, endMs: msOfKey("2026-07-17") + 570 * 60000, allDay: false, location: "Room 4", description: "agenda" }, "UTC");
	eq(body.subject, "Standup", "event body trims the title");
	eq(body.location, { displayName: "Room 4" }, "event body carries the location");
	eq(body.body, { contentType: "text", content: "agenda" }, "event body carries the description as text");
	const bare = graphEventBody({ title: "", startMs: 0, endMs: 0, allDay: false }, "UTC");
	eq(bare.subject, "(no title)", "an empty title still makes a subject");
	eq("location" in bare, false, "omitted fields stay out of the body");
}
// --- meeting essentials: invites, recurrence, RSVP mapping ---
eq(parseAttendeeInput("a@x.com, Bob Lee <b@x.com>; c@x.com"), [{ email: "a@x.com" }, { name: "Bob Lee", email: "b@x.com" }, { email: "c@x.com" }], "attendee input parses bare and named forms");
eq(parseAttendeeInput("not an email, also-not"), [], "tokens without an @ are dropped, not guessed");
eq(parseAttendeeInput(""), [], "empty invite input is empty");
eq(graphRecurrence("none", "2026-07-17"), null, "no repeat, no recurrence");
eq(graphRecurrence("weekly", "2026-07-17"), { pattern: { type: "weekly", interval: 1, daysOfWeek: ["friday"] }, range: { type: "noEnd", startDate: "2026-07-17" } }, "graph weekly repeats on the start's weekday");
eq((graphRecurrence("weekdays", "2026-07-17") as { pattern: { daysOfWeek: string[] } }).pattern.daysOfWeek, ["monday", "tuesday", "wednesday", "thursday", "friday"], "graph weekdays rule");
eq((graphRecurrence("monthly", "2026-07-17") as { pattern: { dayOfMonth: number } }).pattern.dayOfMonth, 17, "graph monthly pins the day of month");
eq(googleRecurrence("weekly", "2026-07-17"), ["RRULE:FREQ=WEEKLY;BYDAY=FR"], "google weekly rrule");
eq(googleRecurrence("weekdays", "2026-07-17"), ["RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR"], "google weekdays rrule");
{
	const startMs = msOfKey("2026-07-17") + 540 * 60000;
	const body = graphEventBody({ title: "Sync", startMs, endMs: startMs + 30 * 60000, allDay: false, attendees: [{ name: "Bob", email: "b@x.com" }], repeat: "weekly" }, "UTC");
	eq(body.attendees, [{ emailAddress: { address: "b@x.com", name: "Bob" }, type: "required" }], "graph body carries invitees");
	eq((body.recurrence as { pattern: { type: string } }).pattern.type, "weekly", "graph body carries the recurrence");
	const g = googleEventBody({ title: "Sync", startMs, endMs: startMs + 30 * 60000, allDay: false, attendees: [{ email: "b@x.com" }], repeat: "daily" }, "UTC");
	eq(g.attendees, [{ email: "b@x.com" }], "google body carries invitees");
	eq(g.recurrence, ["RRULE:FREQ=DAILY"], "google body carries the rrule");
	const noTouch = graphEventBody({ title: "Sync", startMs, endMs: startMs + 30 * 60000, allDay: false }, "UTC");
	eq("attendees" in noTouch, false, "untouched attendees stay out of a patch body");
}
{
	const raw = {
		id: "R",
		subject: "Invite",
		start: { dateTime: "2026-07-17T09:00:00.0000000" },
		end: { dateTime: "2026-07-17T09:30:00.0000000" },
		isOrganizer: false,
		showAs: "free",
		attendees: [{ type: "required", emailAddress: { name: "Power Plugins", address: "s@x.com" } }],
		responseStatus: { response: "tentativelyAccepted" },
	};
	const pc = graphToPC(raw, { sourceId: "m", writable: true });
	eq(pc?.canRsvp, true, "an invite on a writable account can be responded to");
	eq(pc?.myResponse, "tentative", "graph tentativelyAccepted maps to tentative");
	eq(pc?.transparent, true, "showAs free marks the event transparent");
	eq(pc?.attendeeDetail, [{ name: "Power Plugins", email: "s@x.com" }], "attendee detail keeps addresses");
	eq(graphToPC({ ...raw, isOrganizer: true }, { sourceId: "m", writable: true })?.canRsvp, false, "organizers do not RSVP their own meeting");
}
{
	const g = googleToPC(
		{
			id: "gr",
			summary: "Invite",
			start: { dateTime: "2026-07-17T09:00:00Z" },
			end: { dateTime: "2026-07-17T10:00:00Z" },
			transparency: "transparent",
			organizer: { email: "o@x.com" },
			attendees: [{ email: "me@x.com", self: true, responseStatus: "needsAction" }],
		},
		{ sourceId: "g", writable: true }
	);
	eq(g?.canRsvp, true, "a google invite with a self attendee can be responded to");
	eq(g?.myResponse, "none", "needsAction maps to none");
	eq(g?.transparent, true, "google transparency maps");
}

eq(dedupePeople("Power Plugins", ["Power Plugins", "Ana Lee"]), ["Ana Lee"], "the organizer's duplicate attendee row drops");
eq(dedupePeople("power plugins ", ["Power Plugins", "Bob"]), ["Bob"], "dedupe is case and whitespace insensitive");
eq(dedupePeople(undefined, ["Ana"]), ["Ana"], "no organizer, no dedupe");
{
	const raw = { id: "X", subject: "S", start: { dateTime: "2026-07-17T09:00:00.0000000" }, end: { dateTime: "2026-07-17T09:30:00.0000000" }, isOrganizer: false };
	eq(graphToPC(raw, { sourceId: "m", writable: true })?.canEdit, false, "not the organizer, not editable");
	eq(graphToPC({ ...raw, isOrganizer: true }, { sourceId: "m", writable: true })?.canEdit, true, "organizer on a writable source is editable");
	eq(graphToPC({ ...raw, isOrganizer: true }, { sourceId: "m", writable: false })?.canEdit, false, "a read-only connection is never editable");
}

// --- Google mapping ---
{
	const timed = googleToPC(
		{
			id: "g1",
			summary: "Family dinner",
			start: { dateTime: "2026-07-17T09:00:00-04:00" },
			end: { dateTime: "2026-07-17T09:45:00-04:00" },
			organizer: { displayName: "Power Plugins", email: "s@x.com" },
			attendees: [
				{ displayName: "Power Plugins", self: true, responseStatus: "accepted" },
				{ email: "deanna@x.com", responseStatus: "declined" },
				{ displayName: "Projector", resource: true },
			],
			hangoutLink: "https://meet.google.com/abc",
			htmlLink: "https://calendar.google.com/event?eid=1",
			recurringEventId: "g1parent",
		},
		{ sourceId: "g", calendarName: "Steve & Deanna", color: "#2e7d32", writable: true }
	);
	eq(timed?.startMs, Date.UTC(2026, 6, 17, 13, 0, 0), "google timed events parse their RFC3339 offset");
	eq(timed?.attendees, ["Power Plugins", "deanna@x.com"], "google attendees keep people, drop resources");
	eq(timed?.joinUrl, "https://meet.google.com/abc", "hangoutLink is the join url");
	eq(timed?.recurring, true, "a recurringEventId marks the instance recurring");
	eq(timed?.seriesId, "g1parent", "the google series master id rides along");
	eq(timed?.canEdit, true, "writable google calendar makes events editable");
	eq(timed?.declined, false, "someone else declining is not my decline");

	const mine = googleToPC(
		{ id: "g2", summary: "S", start: { dateTime: "2026-07-17T09:00:00Z" }, end: { dateTime: "2026-07-17T10:00:00Z" }, attendees: [{ self: true, email: "me@x.com", responseStatus: "declined" }] },
		{ sourceId: "g" }
	);
	eq(mine?.declined, true, "my own declined response maps");

	const allDay = googleToPC({ id: "g3", summary: "Trip", start: { date: "2026-07-20" }, end: { date: "2026-07-23" } }, { sourceId: "g" });
	eq(allDay && eventDaySpan(allDay), { startKey: "2026-07-20", endKey: "2026-07-22" }, "google all-day end date is exclusive");
	eq(googleToPC({ id: "g4", status: "cancelled", start: { date: "2026-07-20" } }, { sourceId: "g" }), null, "cancelled google events drop");
	const html = googleToPC(
		{ id: "g5", summary: "S", start: { dateTime: "2026-07-17T09:00:00Z" }, end: { dateTime: "2026-07-17T10:00:00Z" }, description: "<b>Agenda</b><br>1. Hello" },
		{ sourceId: "g" }
	);
	eq(html?.description, "Agenda\n1. Hello", "an html google description reads as text");
}
{
	const t = googleTimesBody(msOfKey("2026-07-17") + 540 * 60000, msOfKey("2026-07-17") + 600 * 60000, false, "America/Indiana/Indianapolis");
	eq(t.start, { dateTime: "2026-07-17T09:00:00", timeZone: "America/Indiana/Indianapolis" }, "google timed body carries the zone");
	const ad = googleTimesBody(msOfKey("2026-07-17"), msOfKey("2026-07-18"), true, "UTC");
	eq(ad, { start: { date: "2026-07-17" }, end: { date: "2026-07-18" } }, "google all-day body keeps the exclusive end date");
	const body = googleEventBody({ title: " Dinner ", startMs: msOfKey("2026-07-17") + 540 * 60000, endMs: msOfKey("2026-07-17") + 600 * 60000, allDay: false, location: "Home" }, "UTC");
	eq(body.summary, "Dinner", "google body trims the summary");
	eq(body.location, "Home", "google body carries the location");
	eq("description" in body, false, "omitted description stays out");
}
{
	const payload = { preferred_username: "alex.kim@example.com", name: "Alex Kïm" };
	const b64url = (s: string) => btoa(String.fromCharCode(...new TextEncoder().encode(s))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
	const jwt = `${b64url('{"alg":"none"}')}.${b64url(JSON.stringify(payload))}.sig`;
	eq(decodeJwtPayload(jwt)?.preferred_username, "alex.kim@example.com", "jwt payload decodes");
	eq(decodeJwtPayload(jwt)?.name, "Alex Kïm", "jwt decode survives unicode");
	eq(decodeJwtPayload("garbage"), null, "a non-jwt decodes to null");
}

// --- mergeForSave ---
{
	interface S {
		favs: string[];
		recent: string[];
		token: string;
	}
	const baseline: S = { favs: ["a"], recent: ["x"], token: "t1" };
	const ours: S = { favs: ["a"], recent: ["x", "y"], token: "t1" };
	const disk: S = { favs: ["a", "b"], recent: ["x"], token: "t2" };
	eq(mergeForSave(ours, baseline, disk), { favs: ["a", "b"], recent: ["x", "y"], token: "t2" }, "only our changed keys overwrite disk");
	eq(mergeForSave(ours, baseline, null), ours, "no disk file writes memory as-is");
	eq(mergeForSave({ favs: [] } as unknown as S, { favs: ["a"] } as unknown as S, { favs: ["a", "b"] } as unknown as S), { favs: [] }, "an intentional clear is a change and wins");
	eq(mergeForSave(ours, baseline, { favs: ["a"] } as Partial<S>), { favs: ["a"], recent: ["x", "y"], token: "t1" }, "keys missing on disk keep ours");
}

{
	// A key holding one value per item is a whole vault's worth of settings behind
	// a single name. Changing ONE of them used to publish ALL of them, erasing
	// every item another device had configured since this one last read.
	type M = { map: Record<string, number[]> };
	eq(
		mergeForSave({ map: { A: [2] } } as M, { map: { A: [1] } } as M, { map: { A: [1], B: [9] } } as M),
		{ map: { A: [2], B: [9] } },
		"one entry's change publishes that entry, not the whole map"
	);
	eq(
		mergeForSave({ map: { A: [1] } } as M, { map: { A: [1], B: [9] } } as M, { map: { A: [1], B: [9] } } as M),
		{ map: { A: [1] } },
		"an entry we removed stays removed"
	);
	eq(
		mergeForSave({ map: { A: [1] } } as M, { map: { A: [1] } } as M, { map: { A: [7] } } as M),
		{ map: { A: [7] } },
		"an entry we did not touch takes the disk's"
	);
}

// --- ics.ts: parsing and recurrence expansion ---
import { parseIcsEvents } from "./ics";

const SRC = { sourceId: "feed", calendarName: "Feed", color: "#123456" };
const ics = (...lines: string[]) => ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//pc-test//EN", ...lines, "END:VCALENDAR"].join("\r\n");

{
	const text = ics("BEGIN:VEVENT", "UID:simple-1", "DTSTART:20260717T130000Z", "DTEND:20260717T133000Z", "SUMMARY:Simple", "END:VEVENT");
	const evs = parseIcsEvents(text, msOfKey("2026-07-16"), msOfKey("2026-07-19"), SRC);
	eq(evs.length, 1, "ics: a simple event parses");
	eq(evs[0].startMs, Date.UTC(2026, 6, 17, 13, 0, 0), "ics: a Z time is an absolute instant");
	eq(evs[0].title, "Simple", "ics: summary maps to title");
	eq(evs[0].sourceId, "feed", "ics: source id carries through");
	eq(parseIcsEvents(text, msOfKey("2026-08-01"), msOfKey("2026-08-05"), SRC).length, 0, "ics: events outside the window drop");
}

{
	const text = ics(
		"BEGIN:VEVENT",
		"UID:weekly-1",
		"DTSTART:20260703T130000Z",
		"DTEND:20260703T133000Z",
		"RRULE:FREQ=WEEKLY;COUNT=8",
		"EXDATE:20260724T130000Z",
		"SUMMARY:Weekly sync",
		"END:VEVENT"
	);
	const evs = parseIcsEvents(text, msOfKey("2026-07-15"), msOfKey("2026-08-03"), SRC);
	eq(evs.length, 2, "ics: weekly series expands inside the window, EXDATE removed");
	eq(evs[0].startMs, Date.UTC(2026, 6, 17, 13, 0, 0), "ics: first occurrence is Jul 17");
	eq(evs[1].startMs, Date.UTC(2026, 6, 31, 13, 0, 0), "ics: EXDATE skips Jul 24, next is Jul 31");
	eq(evs[0].recurring, true, "ics: expanded occurrences are marked recurring");
	eq(evs[0].id === evs[1].id, false, "ics: each occurrence gets its own id");
}

{
	const text = ics(
		"BEGIN:VEVENT",
		"UID:daily-1",
		"DTSTART:20260713T100000Z",
		"DTEND:20260713T103000Z",
		"RRULE:FREQ=DAILY;COUNT=5",
		"SUMMARY:Daily",
		"END:VEVENT",
		"BEGIN:VEVENT",
		"UID:daily-1",
		"RECURRENCE-ID:20260715T100000Z",
		"DTSTART:20260715T140000Z",
		"DTEND:20260715T143000Z",
		"SUMMARY:Daily (moved)",
		"END:VEVENT",
		"BEGIN:VEVENT",
		"UID:daily-1",
		"RECURRENCE-ID:20260716T100000Z",
		"DTSTART:20260716T100000Z",
		"DTEND:20260716T103000Z",
		"STATUS:CANCELLED",
		"SUMMARY:Daily",
		"END:VEVENT"
	);
	const evs = parseIcsEvents(text, msOfKey("2026-07-12"), msOfKey("2026-07-19"), SRC);
	eq(evs.length, 4, "ics: 5 dailies minus a canceled override, with a moved one");
	const moved = evs.find((e) => e.title === "Daily (moved)");
	eq(moved?.startMs, Date.UTC(2026, 6, 15, 14, 0, 0), "ics: a moved override renders at its new time");
	eq(evs.some((e) => e.startMs === Date.UTC(2026, 6, 16, 10, 0, 0)), false, "ics: a canceled override leaves a hole");
}

{
	const text = ics(
		"BEGIN:VEVENT",
		"UID:ad-1",
		"DTSTART;VALUE=DATE:20260717",
		"SUMMARY:Holiday",
		"END:VEVENT",
		"BEGIN:VEVENT",
		"UID:ad-2",
		"DTSTART;VALUE=DATE:20260720",
		"DTEND;VALUE=DATE:20260723",
		"SUMMARY:Trip",
		"END:VEVENT"
	);
	const evs = parseIcsEvents(text, msOfKey("2026-07-01"), msOfKey("2026-08-01"), SRC);
	const holiday = evs.find((e) => e.title === "Holiday");
	const trip = evs.find((e) => e.title === "Trip");
	eq(holiday?.allDay, true, "ics: VALUE=DATE is all-day");
	eq(holiday && eventDaySpan(holiday), { startKey: "2026-07-17", endKey: "2026-07-17" }, "ics: DTEND-less date event is one local day");
	eq(trip && eventDaySpan(trip), { startKey: "2026-07-20", endKey: "2026-07-22" }, "ics: date DTEND is exclusive");
}

{
	const text = ics(
		"BEGIN:VTIMEZONE",
		"TZID:Tokyo Standard Time",
		"BEGIN:STANDARD",
		"DTSTART:16010101T000000",
		"TZOFFSETFROM:+0900",
		"TZOFFSETTO:+0900",
		"END:STANDARD",
		"END:VTIMEZONE",
		"BEGIN:VEVENT",
		"UID:tz-1",
		"DTSTART;TZID=Tokyo Standard Time:20260717T090000",
		"DTEND;TZID=Tokyo Standard Time:20260717T100000",
		"SUMMARY:Tokyo meeting",
		"END:VEVENT"
	);
	const evs = parseIcsEvents(text, msOfKey("2026-07-15"), msOfKey("2026-07-19"), SRC);
	eq(evs[0]?.startMs, Date.UTC(2026, 6, 17, 0, 0, 0), "ics: an Outlook-style VTIMEZONE resolves (9am Tokyo is midnight UTC)");
}

{
	const text = ics("BEGIN:VEVENT", "UID:float-1", "DTSTART:20260717T090000", "DTEND:20260717T093000", "SUMMARY:Floating", "END:VEVENT");
	const evs = parseIcsEvents(text, msOfKey("2026-07-16"), msOfKey("2026-07-19"), SRC);
	eq(keyOfMs(evs[0]?.startMs ?? 0), "2026-07-17", "ics: a floating time stays on its wall-clock day");
	eq(minutesOfMs(evs[0]?.startMs ?? 0), 540, "ics: a floating 9:00 renders at local 9:00");
}

{
	const text = ics(
		"BEGIN:VEVENT",
		"UID:meta-1",
		"DTSTART:20260718T150000Z",
		"DTEND:20260718T160000Z",
		"SUMMARY:Review",
		"LOCATION:https://us02web.zoom.us/j/999",
		"ORGANIZER;CN=Ana Lee:mailto:ana@x.com",
		"ATTENDEE;CN=Bob;CUTYPE=INDIVIDUAL:mailto:bob@x.com",
		"ATTENDEE;CUTYPE=RESOURCE;CN=Projector:mailto:proj@x.com",
		"ATTENDEE:mailto:carol@x.com",
		"URL:https://example.com/ev/1",
		"END:VEVENT",
		"BEGIN:VEVENT",
		"UID:canceled-1",
		"DTSTART:20260718T170000Z",
		"DTEND:20260718T180000Z",
		"STATUS:CANCELLED",
		"SUMMARY:Nope",
		"END:VEVENT",
		"BEGIN:VEVENT",
		"UID:dur-1",
		"DTSTART:20260718T120000Z",
		"DURATION:PT45M",
		"SUMMARY:By duration",
		"END:VEVENT"
	);
	const evs = parseIcsEvents(text, msOfKey("2026-07-17"), msOfKey("2026-07-20"), SRC);
	const meta = evs.find((e) => e.title === "Review");
	eq(meta?.organizer, "Ana Lee", "ics: organizer prefers CN");
	eq(meta?.attendees, ["Bob", "carol@x.com"], "ics: attendees keep people, drop resources, fall back to address");
	eq(meta?.joinUrl, "https://us02web.zoom.us/j/999", "ics: a join link is found in the location");
	eq(meta?.url, "https://example.com/ev/1", "ics: the URL property maps");
	eq(evs.some((e) => e.title === "Nope"), false, "ics: a canceled single event drops");
	const dur = evs.find((e) => e.title === "By duration");
	eq(dur && dur.endMs - dur.startMs, 45 * 60000, "ics: DURATION computes the end");
}

// --- mail ---
{
	const m = graphMailToPC(
		{ id: "M1", subject: "  RE: Budget review  ", from: { emailAddress: { name: "Ana Lee", address: "ana@x.com" } }, receivedDateTime: "2026-07-17T12:00:00Z", bodyPreview: "Numbers   attached\nthanks", isRead: false, webLink: "https://outlook.com/1", hasAttachments: true },
		"acct1",
		"steve@example.com"
	);
	eq(m?.from, "Ana Lee", "mail maps the sender name");
	eq(m?.unread, true, "isRead false reads unread");
	eq(m?.preview, "Numbers attached thanks", "preview collapses whitespace");
	eq(m?.receivedMs, Date.UTC(2026, 6, 17, 12, 0, 0), "received time parses absolutely");
	eq(graphMailToPC({ subject: "x" }, "a", "l"), null, "no id, no mail");
}
{
	const f = (id: string, name: string, parentId: string | null, unread = 0, kids = 0) =>
		graphFolderToPC({ id, displayName: name, parentFolderId: parentId ?? undefined, unreadItemCount: unread, childFolderCount: kids })!;
	const folders = [
		f("dr", "Drafts", "root"),
		f("in", "Inbox", "root", 101, 3),
		f("von", "Von", "in", 23),
		f("darwin", "Darwin", "in", 10),
		f("upwork", "Upwork", "in", 0, 1),
		f("good", "Good", "upwork"),
		f("orphan", "Stray", "never-fetched"),
	];
	const tree = orderFolderTree(folders, "in");
	eq(tree.map((t) => t.folder.name), ["Inbox", "Darwin", "Upwork", "Good", "Von", "Drafts", "Stray"], "inbox leads, children alphabetical, orphans surface");
	eq(tree.map((t) => t.depth), [0, 1, 1, 2, 1, 0, 0], "depths follow the tree");
	eq(tree.map((t) => t.expandable), [true, false, true, false, false, false, false], "expandable follows fetched children");
	eq(tree[0].folder.unread, 101, "unread counts ride along");
	const foldedInbox = orderFolderTree(folders, "in", new Set(["in"]));
	eq(foldedInbox.map((t) => t.folder.name), ["Inbox", "Drafts", "Stray"], "a collapsed inbox keeps its row and swallows the subtree");
	eq(foldedInbox[0].expandable, true, "a collapsed folder still shows as expandable");
	const foldedUpwork = orderFolderTree(folders, "in", new Set(["upwork"]));
	eq(foldedUpwork.map((t) => t.folder.name), ["Inbox", "Darwin", "Upwork", "Von", "Drafts", "Stray"], "collapsing a nested folder hides only its children");
	eq(graphFolderToPC({ displayName: "x" }), null, "no id, no folder");
	const subtree = folderSubtreeIds(folders, "in");
	eq([...subtree].sort(), ["darwin", "good", "in", "upwork", "von"], "the inbox subtree reaches nested folders");
	eq(subtree.has("dr"), false, "siblings of the inbox stay out");
	eq(folderSubtreeIds(folders, null).size, 0, "no root, empty subtree");
}
{
	const mk = (id: string, min: number, unread = false): PCMail => ({
		id,
		accountId: "a",
		accountLabel: "l",
		from: "f",
		fromAddress: "f@x.com",
		subject: id,
		preview: "",
		receivedMs: msOfKey("2026-07-17") + min * 60000,
		unread,
	});
	const existing = [mk("c", 300), mk("b", 200), mk("a", 100)];
	const merged = mergeDeltaMessages(existing, [mk("d", 400), mk("b", 200, true)], ["a"]);
	eq(merged.map((m) => m.id), ["d", "c", "b"], "delta merge inserts new, keeps order newest first, drops removed");
	eq(merged[2].unread, true, "a changed message replaces its old self");
	eq(mergeDeltaMessages(existing, [], []).length, 3, "an empty delta changes nothing");
	eq(mergeDeltaMessages(existing, [mk("e", 500)], [], 2).map((m) => m.id), ["e", "c"], "the cap keeps only the newest");
}
{
	const day = 86400000;
	const win = padFetchWindow(1000 * day, 1007 * day);
	eq(win.fromMs, 955 * day, "the fetch window reaches 45 days back");
	eq(win.toMs, 1097 * day, "and 90 days forward");
	const tight = padFetchWindow(1000 * day, 1007 * day, 1, 2);
	eq(tight.fromMs, 999 * day, "custom back padding in days");
	eq(tight.toMs, 1009 * day, "custom forward padding in days");
}
eq(subjectToEventTitle("RE: FW: Budget review"), "Budget review", "reply and forward prefixes strip");
eq(subjectToEventTitle("Budget review"), "Budget review", "plain subjects pass through");
{
	const today = "2026-07-17";
	eq(fmtMailTime(msOfKey("2026-07-17") + 570 * 60000, today, false), "9:30 AM", "today's mail shows a clock");
	eq(fmtMailTime(msOfKey("2026-07-02") + 570 * 60000, today, false), "Jul 2", "this year's mail shows a date");
	eq(fmtMailTime(msOfKey("2025-12-30"), today, false), "Dec 30, 2025", "older mail carries its year");
}

// --- attendee availability ---
{
	const base = msOfKey("2026-07-20") + 540 * 60000; // 9:00
	const runs = parseAvailabilityView("002210 34", base, 30);
	eq(runs, [
		{ startMs: base + 60 * 60000, endMs: base + 120 * 60000, kind: "busy" },
		{ startMs: base + 120 * 60000, endMs: base + 150 * 60000, kind: "tentative" },
		{ startMs: base + 210 * 60000, endMs: base + 240 * 60000, kind: "oof" },
	], "availability digits merge into runs; free and working-elsewhere drop");
	eq(slotConflict(runs, base, base + 30 * 60000), null, "a free slot has no conflict");
	eq(slotConflict(runs, base + 90 * 60000, base + 135 * 60000), "busy", "busy outranks tentative in an overlap");
	eq(slotConflict(runs, base + 120 * 60000, base + 150 * 60000), "tentative", "a tentative-only overlap reads soft");
}

// --- vault-notes source ---
{
	const d = vaultDateSpan("2026-07-17");
	eq(d && eventDaySpan(d), { startKey: "2026-07-17", endKey: "2026-07-17" }, "a bare frontmatter date is one all-day");
	eq(d?.allDay, true, "date-only reads all-day");
	const t = vaultDateSpan("2026-07-17T09:30");
	eq(t && [t.allDay, minutesOfMs(t.startMs), (t.endMs - t.startMs) / 60000], [false, 570, 60], "a datetime reads timed with an hour default");
	const te = vaultDateSpan("2026-07-17T09:30", "2026-07-17T11:00");
	eq(te && (te.endMs - te.startMs) / 60000, 90, "an end datetime sets the real length");
	const span = vaultDateSpan("2026-07-20", "2026-07-22");
	eq(span && eventDaySpan(span), { startKey: "2026-07-20", endKey: "2026-07-22" }, "an end date stretches an all-day span inclusively");
	eq(vaultDateSpan("no date"), null, "no date, no event");
	eq(vaultDateSpan(20260717), null, "a bare number is not a date");
}
eq(fmtVaultDate(msOfKey("2026-07-17") + 570 * 60000, true), "2026-07-17T09:30", "vault datetime round-trips");
eq(fmtVaultDate(msOfKey("2026-07-17") + 570 * 60000, false), "2026-07-17", "vault date-only keeps its character");

// --- week numbers, show-as, reminders ---
eq(isoWeekNum("2026-01-01"), 1, "jan 1 2026 (a thursday) is week 1");
eq(isoWeekNum("2026-07-17"), 29, "mid july is week 29");
eq(isoWeekNum("2027-01-01"), 53, "jan 1 2027 (a friday) belongs to 2026's week 53");
eq(isoWeekNum("2026-12-28"), 53, "the last monday of 2026 opens week 53");
{
	const startMs = msOfKey("2026-07-17") + 540 * 60000;
	const b = graphEventBody({ title: "S", startMs, endMs: startMs + 1800000, allDay: false, showAs: "free" }, "UTC");
	eq(b.showAs, "free", "graph body carries showAs");
	const g = googleEventBody({ title: "S", startMs, endMs: startMs + 1800000, allDay: false, showAs: "free" }, "UTC");
	eq(g.transparency, "transparent", "google free maps to transparent");
	const gt = googleEventBody({ title: "S", startMs, endMs: startMs + 1800000, allDay: false, showAs: "tentative" }, "UTC");
	eq(gt.status, "tentative", "google tentative maps to status");
	eq("showAs" in graphEventBody({ title: "S", startMs, endMs: startMs + 1800000, allDay: false }, "UTC"), false, "no showAs, nothing sent");
}
{
	const ev = evAt("2026-07-17", 600, 30); // starts 10:00
	const at = (min: number) => msOfKey("2026-07-17") + min * 60000;
	const fired = new Set<string>();
	eq(dueReminders([ev], at(594), 5, fired).length, 0, "too early, no reminder");
	eq(dueReminders([ev], at(595), 5, fired).length, 1, "lead time reached fires");
	eq(dueReminders([ev], at(600), 5, fired).length, 1, "start time still fires");
	eq(dueReminders([ev], at(602), 5, fired).length, 0, "well past start stays quiet");
	fired.add(ev.id);
	eq(dueReminders([ev], at(596), 5, fired).length, 0, "a fired reminder never repeats");
	eq(dueReminders([{ ...ev, id: "d", declined: true }], at(596), 5, new Set()).length, 0, "declined meetings stay quiet");
	eq(dueReminders([allDayEv("2026-07-17", 1)], at(596), 5, new Set()).length, 0, "all-day banners stay quiet");
}

// --- availability ---
{
	const evs = [
		evAt("2026-07-20", 600, 60), // 10:00-11:00
		evAt("2026-07-20", 840, 30, { declined: true }), // declined does not block
		evAt("2026-07-20", 900, 30, { transparent: true }), // free does not block
		allDayEv("2026-07-20", 1), // banners do not block
	];
	eq(freeSlotsForDay(evs, "2026-07-20", 540, 1020), [
		{ startMin: 540, endMin: 600 },
		{ startMin: 660, endMin: 1020 },
	], "gaps skip busy, ignore declined, free, and all-day");
	eq(freeSlotsForDay([], "2026-07-20", 540, 1020), [{ startMin: 540, endMin: 1020 }], "an empty day is one full slot");
	eq(freeSlotsForDay([evAt("2026-07-20", 480, 600)], "2026-07-20", 540, 1020), [], "a fully covered window has no slots");
	eq(freeSlotsForDay([evAt("2026-07-20", 600, 40)], "2026-07-20", 540, 660, 30), [{ startMin: 540, endMin: 600 }], "a tail gap under the minimum is not offered");
}
{
	const text = fmtFreeSlots(
		[
			{ key: "2026-07-20", slots: [{ startMin: 540, endMin: 1020 }] },
			{ key: "2026-07-21", slots: [] },
			{ key: "2026-07-22", slots: [{ startMin: 540, endMin: 600 }, { startMin: 780, endMin: 1020 }] },
		],
		540,
		1020,
		false
	);
	eq(text.split("\n")[0], "Mon Jul 20: free", "a fully open day says free");
	eq(text.split("\n")[1], "Tue Jul 21: booked", "a covered day says booked");
	eq(text.split("\n")[2], "Wed Jul 22: 9 AM - 10 AM, 1 PM - 5 PM", "slots list readable ranges");
}
eq(fmtZoneClock(Date.UTC(2026, 6, 17, 13, 30), "America/New_York", false), "9:30 AM", "second zone renders its wall clock");
eq(fmtZoneClock(Date.UTC(2026, 6, 17, 13, 0), "America/New_York", false), "9 AM", "on-the-hour drops :00 in 12h");
eq(fmtZoneClock(Date.UTC(2026, 6, 17, 13, 0), "UTC", true), "13:00", "24h form");
eq(fmtZoneClock(0, "Not/AZone", false), null, "an invalid zone is null, not a throw");

// --- daily-note agenda block ---
eq(dateKeyOf("2026-07-18 Saturday.md"), "2026-07-18", "a daily note's filename yields its day");
eq(dateKeyOf("no date here"), null, "no date is null");
eq(dateKeyOf("2026-13-40"), null, "impossible dates are rejected");
eq(parseAgendaBlock(""), { date: null, days: 1 }, "empty block defaults to one derived day");
eq(parseAgendaBlock("date: today\ndays: 3"), { date: null, days: 3 }, "explicit today stays derived-at-render");
eq(parseAgendaBlock("date: 2026-07-20"), { date: "2026-07-20", days: 1 }, "a fixed date pins the block");
eq(parseAgendaBlock("days: 99\nnonsense line"), { date: null, days: 31 }, "days clamp and junk is ignored");

// --- caldavxml.ts ---
import { buildCalendarQuery, firstTag, hasEmptyTag, parseCalendarData, parseCollections, parseHomeSet, parsePrincipal, resolveHref, toBasicUtc, xmlUnescape } from "./caldavxml";

eq(resolveHref("https://caldav.icloud.com", "/123/principal/"), "https://caldav.icloud.com/123/principal/", "dav: root-relative href takes the origin");
eq(resolveHref("https://x.com/dav/", "cal1/"), "https://x.com/dav/cal1/", "dav: relative href appends to the base");
eq(resolveHref("https://x.com/dav/", "https://other.com/a/"), "https://other.com/a/", "dav: absolute href stays");
eq(toBasicUtc(Date.UTC(2026, 6, 17, 13, 5, 9)), "20260717T130509Z", "dav: basic UTC stamp");
eq(buildCalendarQuery(Date.UTC(2026, 6, 1), Date.UTC(2026, 8, 1)).includes('<c:time-range start="20260701T000000Z" end="20260901T000000Z"/>'), true, "dav: query carries the window");
eq(hasEmptyTag("<d:resourcetype><d:collection/><cal:calendar/></d:resourcetype>", "calendar"), true, "dav: self-closed marker found");
eq(hasEmptyTag("<d:resourcetype><d:collection/></d:resourcetype>", "calendar"), false, "dav: absent marker is false");
eq(firstTag("<a><D:href>/x/</D:href></a>", "href"), "/x/", "dav: local-name match ignores the prefix");
eq(xmlUnescape("A &amp; B&#13;\n&#x41;"), "A & B\r\nA", "dav: entities decode");

{
	const principalXml = `<?xml version="1.0" encoding="UTF-8"?>
<multistatus xmlns="DAV:"><response><href>/</href><propstat><prop>
<current-user-principal><href>/123456/principal/</href></current-user-principal>
</prop><status>HTTP/1.1 200 OK</status></propstat></response></multistatus>`;
	eq(parsePrincipal(principalXml, "https://caldav.icloud.com"), "https://caldav.icloud.com/123456/principal/", "dav: principal parses from a no-prefix multistatus");

	const homeXml = `<D:multistatus xmlns:D="DAV:" xmlns:C="urn:ietf:params:xml:ns:caldav"><D:response><D:href>/123456/principal/</D:href>
<D:propstat><D:prop><C:calendar-home-set><D:href>https://p42-caldav.icloud.com/123456/calendars/</D:href></C:calendar-home-set></D:prop></D:propstat>
</D:response></D:multistatus>`;
	eq(parseHomeSet(homeXml, "https://caldav.icloud.com"), "https://p42-caldav.icloud.com/123456/calendars/", "dav: home set parses across prefixes");
}

{
	const xml = `<d:multistatus xmlns:d="DAV:" xmlns:cal="urn:ietf:params:xml:ns:caldav" xmlns:x1="http://apple.com/ns/ical/">
<d:response><d:href>/remote.php/dav/calendars/steve/</d:href>
<d:propstat><d:prop><d:resourcetype><d:collection/></d:resourcetype></d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response>
<d:response><d:href>/remote.php/dav/calendars/steve/personal/</d:href>
<d:propstat><d:prop><d:displayname>Personal</d:displayname>
<d:resourcetype><d:collection/><cal:calendar/></d:resourcetype>
<x1:calendar-color>#0082C9FF</x1:calendar-color>
<cal:supported-calendar-component-set><cal:comp name="VEVENT"/><cal:comp name="VTODO"/></cal:supported-calendar-component-set>
</d:prop><d:status>HTTP/1.1 200 OK</d:status></d:propstat></d:response>
<d:response><d:href>/remote.php/dav/calendars/steve/tasks/</d:href>
<d:propstat><d:prop><d:displayname>Tasks</d:displayname>
<d:resourcetype><d:collection/><cal:calendar/></d:resourcetype>
<cal:supported-calendar-component-set><cal:comp name="VTODO"/></cal:supported-calendar-component-set>
</d:prop></d:propstat></d:response>
<d:response><d:href>/remote.php/dav/calendars/steve/unnamed/</d:href>
<d:propstat><d:prop><d:displayname/>
<d:resourcetype><d:collection/><cal:calendar/></d:resourcetype>
</d:prop></d:propstat></d:response>
</d:multistatus>`;
	const cols = parseCollections(xml, "https://cloud.example.com/remote.php/dav/calendars/steve/");
	eq(cols.length, 2, "dav: home and task-only collections are excluded");
	eq(cols[0], { href: "https://cloud.example.com/remote.php/dav/calendars/steve/personal/", name: "Personal", color: "#0082c9" }, "dav: a calendar parses with color trimmed to rgb");
	eq(cols[1].name, "unnamed", "dav: a missing displayname falls back to the href segment");
}

{
	const report = `<d:multistatus xmlns:d="DAV:" xmlns:cal="urn:ietf:params:xml:ns:caldav">
<d:response><d:href>/cal/1.ics</d:href><d:propstat><d:prop><d:getetag>"abc"</d:getetag>
<cal:calendar-data>BEGIN:VCALENDAR&#13;
VERSION:2.0&#13;
BEGIN:VEVENT&#13;
UID:r1&#13;
DTSTART:20260717T130000Z&#13;
DTEND:20260717T133000Z&#13;
SUMMARY:A &amp; B&#13;
END:VEVENT&#13;
END:VCALENDAR</cal:calendar-data>
</d:prop></d:propstat></d:response></d:multistatus>`;
	const payloads = parseCalendarData(report);
	eq(payloads.length, 1, "dav: calendar-data extracted");
	const evs = parseIcsEvents(payloads[0], msOfKey("2026-07-16"), msOfKey("2026-07-19"), SRC);
	eq(evs[0]?.title, "A & B", "dav: unescaped payload parses end to end");
}

// --- Graph $batch: chunking and reply parsing ---
{
	const ok = (cond: unknown, msg: string) => eq(!!cond, true, msg);
	const { chunk, GRAPH_BATCH_MAX, buildMessageBatch, parseMessageBatch } = require("./core");

	// -- chunk --
	eq(chunk([1, 2, 3, 4, 5], 2), [[1, 2], [3, 4], [5]], "chunk splits into runs and keeps the remainder");
	eq(chunk([], 20), [], "an empty list yields no chunks");
	eq(chunk([1, 2, 3], 10).length, 1, "a short list is one chunk");
	eq(chunk([1, 2, 3], 0).length, 3, "a zero size cannot make a zero-length chunk (no infinite loop)");
	eq(GRAPH_BATCH_MAX, 20, "the batch ceiling is Graph's documented twenty");

	// -- request building --
	{
		const b = buildMessageBatch(["AAA", "BBB"], "subject,body");
		eq(b.requests.length, 2, "one sub-request per id");
		eq(b.requests[0].id, "0", "sub-requests are keyed by index, not the opaque message id");
		eq(b.requests[0].method, "GET", "reads are GETs");
		ok(b.requests[0].url.includes("/me/messages/AAA?$select=subject,body"), "the id and select ride the sub-url");
		// an id with url-unsafe characters must be encoded
		ok(buildMessageBatch(["a/b+c="], "x").requests[0].url.includes("a%2Fb%2Bc%3D"), "message ids are url-encoded");
	}

	// -- reply parsing --
	const ids = ["AAA", "BBB", "CCC"];
	{
		// responses come back in arbitrary order; they must map back by id
		const json = {
			responses: [
				{ id: "2", status: 200, body: { subject: "third" } },
				{ id: "0", status: 200, body: { subject: "first" } },
				{ id: "1", status: 404, body: { error: {} } },
			],
		};
		const out = parseMessageBatch(json, ids);
		eq(out.ok.size, 2, "successful sub-responses are collected");
		eq(out.ok.get("AAA").subject, "first", "an out-of-order response still maps to the right id");
		eq(out.ok.get("CCC").subject, "third", "the last id resolves correctly");
		eq(out.failed, ["BBB"], "a failed sub-response is reported, not lost");
	}
	{
		// THE ONE THAT MATTERS: a message Graph omits entirely must be reported
		// failed so the caller retries it, never silently dropped
		const out = parseMessageBatch({ responses: [{ id: "0", status: 200, body: { subject: "x" } }] }, ids);
		eq(out.failed.sort(), ["BBB", "CCC"], "ids with no response at all are reported failed");
	}
	{
		const out = parseMessageBatch({ responses: [{ id: "1", status: 429, headers: { "Retry-After": "7" } }] }, ids);
		eq(out.retryAfterMs, 7000, "a throttled sub-response surfaces its Retry-After in ms");
		ok(out.failed.includes("BBB"), "and the throttled id is failed so it will be retried");
	}
	eq(parseMessageBatch(null, ids).failed, ids, "a null reply fails every id");
	eq(parseMessageBatch({ nope: 1 }, ids).failed, ids, "a reply without a responses array fails every id");
	eq(parseMessageBatch({ responses: [] }, ids).failed.sort(), ["AAA", "BBB", "CCC"], "an empty responses array fails every id");
	{
		// a 200 with no body is not a success, there is nothing to cache
		const out = parseMessageBatch({ responses: [{ id: "0", status: 200 }] }, ids);
		ok(out.failed.includes("AAA"), "a 200 carrying no body is treated as a failure");
	}
}

if (failures) {
	console.error(`\n${failures} test(s) failed`);
	process.exit(1);
}
console.log("\nAll tests passed");

// --- the deploy guard ---
// Two sessions building this plugin at once is enough for the second to
// overwrite the first with an older build, silently. The comparison is where a
// bug would disable the guard without failing anything, so it is pinned here.
{
	const { compareVersions, isDowngrade, versionFromManifest } = require("../deploy-guard.mjs");

	eq(compareVersions("1.89.1", "1.89.0") > 0, true, "a later patch sorts after");
	eq(compareVersions("1.89.0", "1.89.1") < 0, true, "and an earlier one before");
	eq(compareVersions("1.89.1", "1.89.1"), 0, "the same version ties");
	// the whole reason this compares numbers: as strings, "1.9.0" sorts after
	// "1.10.0", which is exactly backwards
	eq(compareVersions("1.10.0", "1.9.0") > 0, true, "10 is a later minor than 9, not an earlier one");
	eq(compareVersions("1.88.10", "1.88.9") > 0, true, "and the same holds for the patch");
	eq(compareVersions("2.0.0", "1.99.99") > 0, true, "a major bump outranks everything under it");
	eq(compareVersions("1.89", "1.89.0"), 0, "a missing part counts as zero");
	eq(compareVersions("", ""), 0, "two unreadable versions tie rather than throwing");

	eq(isDowngrade("1.89.1", "1.88.1"), true, "deploying an older build over a newer one is the collision this catches");
	eq(isDowngrade("1.88.1", "1.89.1"), false, "the ordinary direction is not");
	eq(isDowngrade("1.89.1", "1.89.1"), false, "and neither is redeploying the same version, which is what developing looks like");
	eq(isDowngrade(null, "1.89.1"), false, "a vault with nothing installed has nothing to lose");
	eq(isDowngrade("", "1.89.1"), false, "nor one whose version could not be read");

	eq(versionFromManifest("{ not json"), null, "a manifest too broken to parse names no version");
	eq(versionFromManifest("{}"), null, "and neither does one with no version key");
	eq(versionFromManifest('{"version":"1.2.3"}'), "1.2.3", "otherwise the version is read off it");
	eq(versionFromManifest('{"version":"  "}'), null, "a blank version is no version");
}
