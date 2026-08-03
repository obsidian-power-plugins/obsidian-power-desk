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
	isSystemFolder,
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
	attachmentBadge,
	fmtAttachmentSize,
	mimeForExtension,
	groupThreads,
	printableHtml,
	printableTableHtml,
	printableAgendaHtml,
	printableMonthHtml,
	pageRule,
	scaledPt,
	PRINT_SCALES,
	signatureFor,
	extractInlineImages,
	migrateSignature,
	newArrivals,
	arrivalSummary,
	automatedSender,
	sectionOf,
	splitSections,
	toGraphDateTime,
	fromGraphDateTime,
	categoryColor,
	CATEGORY_COLOR_NAMES,
	toggleCategory,
	replaceCategory,
	buildCategoryPatchBatch,
	buildReadPatchBatch,
	parseWriteBatch,
	groupShortcuts,
	defaultShortcutLabel,
	Shortcut,
	buildJournal,
	journalMarkdown,
	matchEvents,
	eventQueryIsEmpty,
	EventQuery,
	parseUnsubscribe,
	unsubscribePlan,
	ruleToEdit,
	ruleToBody,
	ruleHasUnknownParts,
	ruleSummary,
	EMPTY_RULE,
	tokenize,
	buildQuery,
	graphSearchText,
	passesLocalFilters,
	parseQuery,
	buildIndex,
	searchIndex,
	IndexDoc,
	whenPresets,
	rankContacts,
	mergePeople,
	matchContacts,
	currentAddressFragment,
	applyAddressChoice,
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

// --- attachment badges ---
eq(attachmentBadge("Invoice-T-258926.pdf", "application/pdf").label, "PDF", "the extension is the label");
eq(attachmentBadge("Invoice.pdf", "application/pdf").color, attachmentBadge("other.PDF", "").color, "the extension match ignores case");
eq(attachmentBadge("Q3.xlsx", "").color === attachmentBadge("Q3.docx", "").color, false, "Excel and Word take different colors");
eq(attachmentBadge("deck.pptx", "").label, "PPTX", "a four-character extension survives whole");
eq(attachmentBadge("archive.tar.gz", "").label, "GZ", "the last extension is the one that counts");
eq(attachmentBadge("image001", "image/png").label, "IMG", "no extension falls back to the MIME family");
eq(attachmentBadge("mystery", "application/x-nonsense").label, "FILE", "an unreadable type still gets a badge");
eq(attachmentBadge("mystery", "").color, attachmentBadge("thing.qqq", "").color, "unknowns share the neutral gray");
eq(attachmentBadge("Report.docx", "").label.length <= 4, true, "labels never exceed four characters");
eq(fmtAttachmentSize(0), "1 KB", "a file that exists never reads as 0 KB");
eq(fmtAttachmentSize(72704), "71 KB", "kilobytes round to whole numbers");
eq(fmtAttachmentSize(1048576), "1.0 MB", "a megabyte crosses to one decimal");
eq(fmtAttachmentSize(5872025), "5.6 MB", "megabytes keep one decimal");
eq(fmtAttachmentSize(2147483648), "2.0 GB", "gigabytes have their own unit");
eq(mimeForExtension("pdf"), "application/pdf", "an extension resolves to its content type");
eq(mimeForExtension("PDF"), "application/pdf", "whatever the case");
eq(mimeForExtension(".png"), "image/png", "a leading dot is tolerated");
eq(mimeForExtension("xlsx"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "the Office types are the long ones");
eq(mimeForExtension("qqq"), "application/octet-stream", "an unknown extension falls back to bytes");
eq(mimeForExtension(""), "application/octet-stream", "and so does no extension at all");

// --- conversation threading ---
{
	const mail = (id: string, conv: string | undefined, ms: number, from: string, extra: Partial<PCMail> = {}, accountId = "a1"): PCMail => ({
		id,
		accountId,
		accountLabel: "Work",
		from,
		fromAddress: `${from.toLowerCase()}@x.com`,
		subject: "Re: Tank Management",
		preview: "",
		receivedMs: ms,
		unread: false,
		conversationId: conv,
		...extra,
	});
	const list = [
		mail("m3", "c1", 3000, "Deanna", { unread: true }),
		mail("m2", "c2", 2000, "Chris"),
		mail("m1", "c1", 1000, "Steve", { hasAttachments: true }),
		mail("m0", "c1", 500, "Deanna", { flagged: true, priority: true }),
	];
	const t = groupThreads(list);
	eq(t.length, 2, "four messages over two conversations give two threads");
	eq(t[0].messages.map((m) => m.id).join(","), "m3,m1,m0", "a thread holds its messages newest first");
	eq(t[0].latest.id, "m3", "the newest message heads the thread");
	eq(t[1].messages.length, 1, "a conversation of one is still a thread");
	eq(t[0].unread, 1, "the thread counts its unread");
	eq(t[0].senders.join(","), "Deanna,Steve", "senders are distinct, newest first");
	eq(t[0].hasAttachments, true, "an attachment anywhere in the thread rolls up");
	eq(t[0].flagged && t[0].priority, true, "so do the flag and the importance mark");
	eq(t[1].hasAttachments || t[1].flagged || t[1].priority, false, "and a thread with none of them rolls up nothing");

	// order follows the list it was handed, so whatever sorted it still wins
	eq(t.map((x) => x.latest.id).join(","), "m3,m2", "threads sit where their first message sat");
	const asc = groupThreads([...list].reverse());
	eq(asc.map((x) => x.latest.id).join(","), "m3,m2", "an oldest-first list still names the newest as latest");
	eq(asc[0].messages[0].id, "m3", "and still orders the thread newest first");

	// a cache written before threading has no conversation ids at all
	const legacy = groupThreads([mail("x1", undefined, 2000, "Ana"), mail("x2", undefined, 1000, "Bo")]);
	eq(legacy.length, 2, "messages with no conversation id are threads of one");

	// conversation ids are only unique inside a mailbox
	const cross = groupThreads([mail("y1", "shared", 2000, "Ana", {}, "a1"), mail("y2", "shared", 1000, "Bo", {}, "a2")]);
	eq(cross.length, 2, "the same conversation id in two accounts does not braid them together");

	eq(groupThreads([]).length, 0, "an empty list gives no threads");
}

// --- automatic replies: the date round trip ---
{
	const ms = Date.UTC(2026, 7, 5, 14, 30, 0);
	const g = toGraphDateTime(ms);
	eq(g.timeZone, "UTC", "times go out absolute, in the one zone spelling nothing argues with");
	eq(g.dateTime, "2026-08-05T14:30:00", "and as a plain wall time with no zone suffix");
	eq(fromGraphDateTime(g), ms, "a UTC value round trips exactly");
	eq(fromGraphDateTime({ dateTime: "2026-08-05T14:30:00.0000000", timeZone: "UTC" }), ms, "Graph's seven-digit fraction parses");
	eq(fromGraphDateTime({ dateTime: "2026-08-05T14:30:00Z", timeZone: "UTC" }), ms, "and so does one that already carries a Z");
	eq(fromGraphDateTime({ dateTime: "2026-08-05T14:30:00", timeZone: "GMT" }), ms, "GMT is read as UTC");

	// a non-UTC zone is read as this machine's wall time, the best guess left
	const local = fromGraphDateTime({ dateTime: "2026-08-05T14:30:00", timeZone: "Central Standard Time" });
	eq(local, new Date(2026, 7, 5, 14, 30, 0).getTime(), "an unrecognized zone reads as local wall time");

	eq(fromGraphDateTime(undefined), null, "nothing parses to nothing");
	eq(fromGraphDateTime({ dateTime: "", timeZone: "UTC" }), null, "and so does an empty value");
	eq(fromGraphDateTime({ dateTime: "not a date", timeZone: "UTC" }), null, "an unreadable value is null, not NaN");
	// the trip a user actually takes: pick a local time, store it, read it back
	const picked = new Date(2026, 11, 24, 17, 0, 0).getTime();
	eq(fromGraphDateTime(toGraphDateTime(picked)), picked, "a local time survives the round trip through UTC");
}

// --- categories ---
{
	eq(categoryColor("preset0"), "#e74856", "the first preset is Outlook's red");
	eq(categoryColor("preset7"), "#0078d4", "and the eighth its blue");
	eq(categoryColor("preset24"), "#93003f", "the last preset resolves");
	eq(categoryColor("PRESET0"), "#e74856", "the slot name is read whatever the case");
	eq(categoryColor("preset25"), "#8a8886", "a slot past the end falls back to gray");
	eq(categoryColor("none"), "#8a8886", "a category with no color set reads gray");
	eq(categoryColor(""), "#8a8886", "and so does nothing at all");
	eq(new Set([0, 7, 12, 19, 24].map((n) => categoryColor(`preset${n}`))).size, 5, "the presets are distinct colors");
	eq(CATEGORY_COLOR_NAMES.length, 25, "every preset has a name to show in a picker");
	eq(CATEGORY_COLOR_NAMES[0], "Red", "and the names line up with Outlook's own order");
	eq(CATEGORY_COLOR_NAMES[7], "Blue", "at both ends of the light half");
	eq(new Set(CATEGORY_COLOR_NAMES).size, 25, "no two colors share a name");

	eq(toggleCategory(undefined, "Red"), ["Red"], "toggling onto nothing starts the list");
	eq(toggleCategory([], "Red"), ["Red"], "and onto an empty list too");
	eq(toggleCategory(["Red"], "Red"), [], "toggling the only one off empties it");
	eq(toggleCategory(["Red", "Blue"], "Red"), ["Blue"], "removing one keeps the rest");
	eq(toggleCategory(["Red", "Blue"], "Green"), ["Red", "Blue", "Green"], "adding appends, keeping the mailbox's order");
	eq(toggleCategory(["Red"], "red"), [], "the match ignores case, as Outlook does");
	eq(toggleCategory(["Red", "Blue"], "BLUE"), ["Red"], "whichever way the case falls");
}

// --- replacing a category everywhere ---
{
	eq(replaceCategory(["Old"], "Old", "New"), ["New"], "the named one is replaced");
	eq(replaceCategory(["Red", "Old", "Blue"], "Old", "New"), ["Red", "New", "Blue"], "and keeps its place among the others");
	eq(replaceCategory(["old"], "Old", "New"), ["New"], "the match ignores case, as Outlook does");
	eq(replaceCategory(["Red"], "Old", "New"), ["Red"], "a message without it is left alone");
	eq(replaceCategory(undefined, "Old", "New"), [], "and one with no categories stays empty");
	// the case that would otherwise show a doubled label in Outlook
	eq(replaceCategory(["Old", "New"], "Old", "New"), ["New"], "a message already carrying both does not end up with two");
	eq(replaceCategory(["New", "Old"], "Old", "New"), ["New"], "whichever order they were in");

	const items = [
		{ id: "a", categories: ["New"] },
		{ id: "b", categories: ["Red", "New"] },
	];
	const batch = buildCategoryPatchBatch(items);
	eq(batch.requests.length, 2, "one sub-request per message");
	eq(batch.requests[0].method, "PATCH", "written as a PATCH");
	eq(batch.requests[0].id, "0", "keyed by index, since Graph ids are too long for the field");
	eq(batch.requests[0].headers?.["Content-Type"], "application/json", "each sub-request carries its own content type");
	eq((batch.requests[1].body as { categories: string[] }).categories, ["Red", "New"], "and its own new category list");
	eq(buildCategoryPatchBatch([]).requests.length, 0, "nothing to write is an empty batch");

	const readBatch = buildReadPatchBatch(["m1", "m2"], true);
	eq(readBatch.requests.length, 2, "one sub-request per message to mark");
	eq(readBatch.requests[0].method, "PATCH", "written as a PATCH");
	eq((readBatch.requests[0].body as { isRead: boolean }).isRead, true, "asking for read");
	eq((buildReadPatchBatch(["m1"], false).requests[0].body as { isRead: boolean }).isRead, false, "and unread when that is what was asked");
	eq(buildReadPatchBatch([], true).requests.length, 0, "nothing to mark is an empty batch");

	const ids = ["a", "b", "c"];
	eq(parseWriteBatch({ responses: [{ id: "0", status: 200 }, { id: "1", status: 200 }, { id: "2", status: 200 }] }, ids).ok, ids, "all happy is all ok");
	eq(parseWriteBatch({ responses: [{ id: "0", status: 200 }, { id: "1", status: 403 }] }, ids).failed, ["b", "c"], "a refusal and a missing answer both count as failed");
	eq(parseWriteBatch(null, ids).failed, ids, "an unreadable reply fails everything, rather than claiming success");
	eq(parseWriteBatch({}, ids).failed, ids, "and so does one with no responses at all");
}

// --- shortcuts ---
{
	const sc = (id: string, group: string, label: string): Shortcut => ({ id, group, label, kind: "url", target: "https://x.com" });
	const groups = groupShortcuts([sc("1", "Work", "a"), sc("2", "", "b"), sc("3", "Work", "c"), sc("4", "Home", "d")]);
	eq(groups.map((g) => g.name).join(","), ",Work,Home", "the unnamed group leads, the rest keep the order they appear in");
	eq(groups.find((g) => g.name === "Work")?.items.length, 2, "members of a group come together");
	eq(groupShortcuts([]).length, 0, "nothing groups into nothing");
	eq(groupShortcuts([sc("1", "  Work  ", "a")])[0].name, "Work", "a group name is trimmed");

	eq(defaultShortcutLabel("note", "Personal/Apple/Apple Leadership.md"), "Apple Leadership", "a note is named by its file");
	eq(defaultShortcutLabel("url", "https://irely.com/"), "irely.com", "a link drops the scheme and the trailing slash");
	eq(defaultShortcutLabel("search", "from:deanna is:unread"), "from:deanna is:unread", "a search is named by itself");
	eq(defaultShortcutLabel("url", `https://x.com/${"a".repeat(200)}`).length <= 60, true, "a very long label is cut to fit a list");
}

// --- journal ---
{
	const day = "2026-08-05";
	const at = (h: number) => msOfKey(day) + h * 3600000;
	const mail = (id: string, from: string, subject: string, ms: number): PCMail => ({
		id,
		accountId: "a",
		accountLabel: "l",
		from,
		fromAddress: `${id}@x.com`,
		subject,
		preview: "",
		receivedMs: ms,
		unread: false,
	});
	const evs: PCEvent[] = [
		{ id: "e1", sourceId: "c", title: "Standup", startMs: at(9), endMs: at(9.5), allDay: false, attendees: ["Von", "Deanna"] },
		{ id: "e2", sourceId: "c", title: "Skipped", startMs: at(11), endMs: at(12), allDay: false, declined: true },
		{ id: "e3", sourceId: "c", title: "Holiday", startMs: msOfKey(day), endMs: msOfKey("2026-08-06"), allDay: true },
		{ id: "e4", sourceId: "c", title: "Tomorrow", startMs: msOfKey("2026-08-06") + at(0) - msOfKey(day), endMs: msOfKey("2026-08-07"), allDay: false },
	];
	const got = [mail("r1", "Chris Tate", "Financials", at(8)), mail("r2", "Late", "Not today", msOfKey("2026-08-06") + 3600000)];
	const put = [mail("s1", "Steve Palm", "Re: Financials", at(10))];
	const j = buildJournal(day, evs, got, put, [{ title: "Kore notes", changedMs: at(14) }, { title: "Old", changedMs: msOfKey("2026-08-01") }], false);

	eq(j.meetings.map((m) => m.title).join(","), "Holiday,Standup", "meetings are listed, a declined one is not");
	eq(j.meetings.find((m) => m.title === "Holiday")?.when, "All day", "an all-day event has no time because it did not happen at one");
	eq(j.meetings.find((m) => m.title === "Standup")?.who, "Von, Deanna", "attendees ride along");
	eq(j.received.map((m) => m.subject).join(","), "Financials", "only that day's mail counts");
	eq(j.sent.map((m) => m.subject).join(","), "Re: Financials", "and what was sent is its own list");
	eq(j.notes.map((n) => n.title).join(","), "Kore notes", "notes touched that day, and no others");

	const md = journalMarkdown(j);
	eq(md.startsWith("## Wednesday, Aug 5"), true, "the markdown leads with the day");
	eq(md.includes("### Meetings"), true, "and has a section per kind");
	eq(md.includes("### Sent"), true, "including what went out");
	eq(md.includes("- **9 AM** Standup with Von, Deanna"), true, "a meeting line reads as one");
	// a heading with nothing under it is worse than no heading
	const bare = journalMarkdown(buildJournal(day, [], [], [], [], false));
	eq(bare.includes("### Meetings"), false, "an empty section is left out");
	eq(bare.includes("_Nothing recorded._"), true, "and a day with nothing in it says so");
	eq(journalMarkdown(buildJournal(day, evs, [], [], [], false)).includes("### Received"), false, "sections appear only when they have something");
}

// --- calendar search ---
{
	const ev = (id: string, title: string, extra: Partial<PCEvent> = {}): PCEvent => ({
		id,
		sourceId: "cal1",
		title,
		startMs: msOfKey("2026-08-05") + 9 * 3600000,
		endMs: msOfKey("2026-08-05") + 10 * 3600000,
		allDay: false,
		...extra,
	});
	const empty: EventQuery = { words: "", title: "", people: "", location: "", calendar: "", onlineOnly: false, allDayOnly: false, withPeopleOnly: false };
	// distinct start times, so the ordering assertions test the sort rather
	// than which way ties happen to fall
	const events = [
		ev("a", "Quarterly Review", { location: "Room 4", attendees: ["Deanna Palm"], organizer: "Steve Palm", startMs: msOfKey("2026-08-06") }),
		ev("b", "Standup", { joinUrl: "https://teams.microsoft.com/x", attendees: ["Von Pasion"], startMs: msOfKey("2026-08-05") }),
		ev("c", "Holiday", { allDay: true, sourceId: "cal2", startMs: msOfKey("2026-08-04") }),
		ev("d", "Review of budgets", { description: "quarterly numbers", startMs: msOfKey("2026-08-07") }),
	];
	const ids = (q: Partial<EventQuery>) => matchEvents(events, { ...empty, ...q }).map((e) => e.id).join(",");

	eq(ids({ words: "review" }), "a,d", "a word search reaches title and description, soonest first");
	eq(ids({ title: "review" }), "a,d", "a title search matches only titles");
	eq(ids({ title: "quarterly review" }), "a", "every word of a title search must match");
	eq(ids({ words: "quarterly" }), "a,d", "but the same words against everything reach the description too");
	eq(ids({ people: "deanna" }), "a", "people matches attendees");
	eq(ids({ people: "steve" }), "a", "and the organizer");
	eq(ids({ location: "room" }), "a", "location matches the place");
	eq(ids({ calendar: "cal2" }), "c", "a calendar filter narrows to one source");
	eq(ids({ onlineOnly: true }), "b", "online only keeps what has a join link");
	eq(ids({ allDayOnly: true }), "c", "all-day only keeps all-day events");
	eq(ids({ withPeopleOnly: true }), "b,a", "with-people only keeps the ones that have any");
	eq(ids({ words: "REVIEW" }), "a,d", "matching ignores case");
	eq(ids({ words: "review", onlineOnly: true }), "", "filters combine, and combining can find nothing");
	eq(ids({}), "c,b,a,d", "asking nothing matches everything, in time order");
	eq(ids({ words: "nothinglikethis" }), "", "a word matching nothing gives nothing");

	eq(eventQueryIsEmpty(empty), true, "an untouched dialog asks nothing");
	eq(eventQueryIsEmpty({ ...empty, words: "  " }), true, "and neither does one with only spaces");
	eq(eventQueryIsEmpty({ ...empty, words: "x" }), false, "a word is a question");
	eq(eventQueryIsEmpty({ ...empty, onlineOnly: true }), false, "and so is a toggle on its own");
	eq(eventQueryIsEmpty({ ...empty, calendar: "cal1" }), false, "and so is picking a calendar");
}

// --- unsubscribe headers ---
{
	const h = (name: string, value: string) => ({ name, value });
	eq(parseUnsubscribe(undefined), null, "no headers, no unsubscribe");
	eq(parseUnsubscribe([]), null, "and neither does an empty set");
	eq(parseUnsubscribe([h("Subject", "hi")]), null, "an unrelated header is not one");
	eq(parseUnsubscribe([h("List-Unsubscribe", "   ")]), null, "an empty header is not one either");

	const web = parseUnsubscribe([h("List-Unsubscribe", "<https://x.com/u?id=1>")])!;
	eq(web.webUrl, "https://x.com/u?id=1", "an https link is read");
	eq(web.oneClickUrl, undefined, "but it is not one-click without the header that says so");

	const oc = parseUnsubscribe([h("List-Unsubscribe", "<https://x.com/u?id=1>"), h("List-Unsubscribe-Post", "List-Unsubscribe=One-Click")])!;
	eq(oc.oneClickUrl, "https://x.com/u?id=1", "the post header makes it one-click");
	eq(unsubscribePlan(oc).kind, "post", "and one-click is what gets used");

	const both = parseUnsubscribe([h("list-unsubscribe", "<mailto:leave@x.com?subject=stop&body=please>, <https://x.com/u>")])!;
	eq(both.webUrl, "https://x.com/u", "a header can carry both kinds");
	eq(both.mailto?.to, "leave@x.com", "and the address parses out");
	eq(both.mailto?.subject, "stop", "with its subject");
	eq(both.mailto?.body, "please", "and its body");
	eq(unsubscribePlan(both).kind, "open", "a link is preferred over sending mail");

	const mailOnly = parseUnsubscribe([h("List-Unsubscribe", "<mailto:leave@x.com>")])!;
	eq(mailOnly.mailto?.subject, "unsubscribe", "a bare mailto gets a sensible subject");
	eq(unsubscribePlan(mailOnly).kind, "mail", "and mail is the only way left");
	eq(unsubscribePlan(mailOnly).target, "leave@x.com", "the plan names where it goes");

	eq(parseUnsubscribe([h("List-Unsubscribe", "<mailto:not-an-address>")]), null, "a mailto with no address is no way out");
	eq(parseUnsubscribe([h("List-Unsubscribe", "https://x.com/u")]), null, "a URI outside angle brackets is not the format");
	eq(parseUnsubscribe([h("List-Unsubscribe", "<ftp://x.com/u>")]), null, "and an unusable scheme is ignored");
	eq(parseUnsubscribe([h("LIST-UNSUBSCRIBE", "<https://x.com/u>")])!.webUrl, "https://x.com/u", "header names match whatever the case");
	// the first of each kind wins, so two links do not fight
	eq(parseUnsubscribe([h("List-Unsubscribe", "<https://a.com/1>, <https://b.com/2>")])!.webUrl, "https://a.com/1", "the first link is the one offered");
}

// --- inbox rules ---
{
	const stored = {
		displayName: "Jira to JIRA",
		isEnabled: true,
		conditions: { senderContains: ["jira@", "atlassian"], subjectContains: ["[JIRA]"] },
		actions: { moveToFolder: "f-jira", markAsRead: true },
		exceptions: {},
	};
	const e = ruleToEdit(stored);
	eq(e.name, "Jira to JIRA", "a stored rule reads back its name");
	eq(e.fromContains, "jira@, atlassian", "a condition list reads back comma separated");
	eq(e.subjectContains, "[JIRA]", "and a single value reads back bare");
	eq(e.moveToFolderId, "f-jira", "the destination folder reads back");
	eq(e.markAsRead && e.enabled, true, "so do the flags");
	eq(e.deleteIt || e.hasAttachments || e.highImportance, false, "and what was not set stays unset");

	// the round trip has to be lossless for what it manages
	const body = ruleToBody(e, stored);
	eq((body.conditions as Record<string, unknown>).senderContains, ["jira@", "atlassian"], "editing and saving preserves a list");
	eq((body.actions as Record<string, unknown>).moveToFolder, "f-jira", "and the destination");
	eq(body.displayName, "Jira to JIRA", "and the name");

	// the part that matters: Outlook can set things this editor never shows,
	// and a PATCH replaces conditions and actions whole
	const richer = {
		displayName: "From Outlook",
		isEnabled: true,
		conditions: { senderContains: ["boss@"], fromAddresses: [{ emailAddress: { address: "boss@x.com" } }], sensitivity: "private" },
		actions: { moveToFolder: "f1", forwardTo: [{ emailAddress: { address: "me@x.com" } }] },
		exceptions: {},
	};
	const kept = ruleToBody(ruleToEdit(richer), richer);
	const kc = kept.conditions as Record<string, unknown>;
	const ka = kept.actions as Record<string, unknown>;
	eq(!!kc.fromAddresses, true, "a condition this editor cannot show survives the save");
	eq(kc.sensitivity, "private", "and so does another one");
	eq(!!ka.forwardTo, true, "and an action it cannot show survives too");
	eq(kc.senderContains, ["boss@"], "while the managed condition is still written");

	// clearing a managed field removes its key rather than writing an empty list
	const cleared = ruleToBody({ ...ruleToEdit(stored), fromContains: "  ", markAsRead: false }, stored);
	eq("senderContains" in (cleared.conditions as Record<string, unknown>), false, "clearing a condition deletes the key");
	eq("markAsRead" in (cleared.actions as Record<string, unknown>), false, "and turning an action off deletes it too");
	eq("subjectContains" in (cleared.conditions as Record<string, unknown>), true, "while the ones still set stay");

	eq(ruleToBody({ ...EMPTY_RULE, name: "  " }).displayName, "Untitled rule", "a rule with no name still gets one");
	eq(Object.keys(ruleToBody(EMPTY_RULE).conditions as Record<string, unknown>).length, 0, "an empty rule writes no conditions");

	// knowing when there is more than this editor shows
	eq(ruleHasUnknownParts(stored), false, "a rule made here has nothing hidden");
	eq(ruleHasUnknownParts(richer), true, "one from Outlook may have");
	eq(ruleHasUnknownParts({ conditions: {}, actions: {}, exceptions: { subjectContains: ["x"] } }), true, "an exception counts as hidden, since none are shown");

	// the one-line summary
	eq(ruleSummary(e, "JIRA"), "When from jira@, atlassian and subject has [JIRA]: move to JIRA, mark read", "a rule reads as a sentence");
	eq(ruleSummary(EMPTY_RULE), "Does nothing yet", "an empty rule says so");
	eq(ruleSummary({ ...EMPTY_RULE, markAsRead: true }), "Everything: mark read", "a rule with no condition applies to everything");
	eq(ruleSummary({ ...EMPTY_RULE, fromContains: "x@y" }), "When from x@y: nothing yet", "a rule with no action says that too");
	eq(ruleSummary({ ...EMPTY_RULE, deleteIt: true, hasAttachments: true }), "When has an attachment: delete", "conditions and actions both read plainly");
}

// --- the local search index ---
{
	eq(tokenize("Invoice T-258926").join(","), "invoice,t-258926,t,258926", "a hyphenated id keeps its whole form and its parts");
	eq(tokenize("steve.palm@irely.com").join(","), "steve.palm@irely.com,steve,palm,irely,com", "an address is findable whole or by any piece");
	eq(tokenize("  Hello,   world!  ").join(","), "hello,world", "punctuation and runs of space fall away");
	eq(tokenize("").length, 0, "nothing tokenizes to nothing");
	eq(tokenize("...---...").length, 0, "and so does punctuation on its own");
	eq(tokenize("RE: Kore").join(","), "re,kore", "a trailing colon is not part of the word");

	const doc = (id: string, subject: string, from: string, body: string, ms: number, extra: Partial<IndexDoc> = {}): IndexDoc => ({
		id,
		subject,
		from,
		body,
		ms,
		unread: false,
		flagged: false,
		hasAttachments: false,
		...extra,
	});
	const idx = buildIndex([
		doc("1", "Invoice T-258926", "Pulseway <support@pulseway.com>", "Thank you for your payment", 500, { hasAttachments: true }),
		doc("2", "Kore Tank Management", "Deanna Palm <deanna@irely.com>", "Can you look at the invoice line", 400, { unread: true }),
		doc("3", "Nirvana Review", "Chris Tate <chris@irely.com>", "August numbers attached", 300, { flagged: true }),
		doc("4", "Penetration Test", "Craig Cannon <craig@irely.com>", "report is ready", 200),
	]);

	const ids = (q: string) => searchIndex(idx, q).join(",");
	eq(ids("invoice"), "1,2", "a subject hit outranks a body hit");
	eq(ids("inv"), "1,2", "a prefix matches, so results narrow as you type");
	eq(ids("258926"), "1", "a ticket number is findable by its digits");
	eq(ids("t-258926"), "1", "and by its whole form");
	eq(ids("palm"), "2", "an address is findable by a name inside it");
	eq(ids("invoice kore"), "2", "two words both have to match");
	eq(ids("invoice zzz"), "", "and a word matching nothing gives nothing");
	eq(ids(""), "", "an empty query matches nothing rather than everything");
	eq(ids("   "), "", "and neither does whitespace");

	// fields and flags
	eq(ids("from:deanna"), "2", "from: matches the sender");
	eq(ids("from:irely.com"), "2,3,4", "including by domain, newest first");
	eq(ids("subject:review"), "3", "subject: matches only the subject");
	eq(ids("subject:payment"), "", "so a body word does not satisfy it");
	eq(ids("is:unread"), "2", "is:unread filters");
	eq(ids("is:read"), "1,3,4", "and so does is:read");
	eq(ids("is:flagged"), "3", "is:flagged filters");
	eq(ids("has:attachment"), "1", "has:attachment filters");
	eq(ids("has:attachments"), "1", "in either spelling");
	eq(ids("invoice is:unread"), "2", "a filter narrows a word search");
	eq(ids("from:irely.com is:flagged"), "3", "and filters combine");

	// phrases
	eq(ids('"tank management"'), "2", "a quoted phrase matches as a run");
	eq(ids('"management tank"'), "", "and not in the wrong order");
	eq(ids('"thank you for your payment"'), "1", "phrases reach into bodies");
	eq(ids('""'), "", "an empty phrase asks for nothing");

	// date bounds
	const dated = buildIndex([
		doc("j1", "Jan", "a@x.com", "", msOfKey("2026-01-15")),
		doc("j2", "Jun", "a@x.com", "", msOfKey("2026-06-15")),
		doc("j3", "Dec", "a@x.com", "", msOfKey("2026-12-15")),
	]);
	const dids = (q: string) => searchIndex(dated, q).join(",");
	eq(dids("after:2026-06-01"), "j3,j2", "after: keeps what came later, newest first");
	eq(dids("before:2026-06-30"), "j2,j1", "before: keeps what came earlier");
	eq(dids("after:2026-06-01 before:2026-06-30"), "j2", "the two together make a range");
	// "before Friday" said out loud includes Friday
	eq(dids("before:2026-06-15"), "j2,j1", "before: includes the whole of that day, not just its midnight");
	eq(dids("after:2026-06-15"), "j3,j2", "and after: starts at that day's beginning");
	eq(dids("after:nonsense"), "", "an unreadable date narrows to nothing rather than being ignored");
	eq(dids("from:a@x.com after:2026-06-01"), "j3,j2", "a date narrows a field search");

	// the dialog builds the same text you could type
	eq(buildQuery({ words: "invoice" }), "invoice", "words pass through");
	eq(buildQuery({ from: "steve, deanna" }), "from:steve from:deanna", "each sender becomes its own term");
	eq(buildQuery({ subject: "tank management" }), "subject:tank subject:management", "and each subject word its own");
	eq(buildQuery({ phrase: "tank management" }), '"tank management"', "a phrase is quoted");
	eq(buildQuery({ phrase: 'a "quoted" thing' }), '"a quoted thing"', "inner quotes are dropped rather than breaking the query");
	eq(buildQuery({ unread: true, flagged: true, attachments: true }), "is:unread is:flagged has:attachment", "the toggles become filters");
	eq(buildQuery({ after: "2026-01-01", before: "2026-12-31" }), "after:2026-01-01 before:2026-12-31", "the dates become bounds");
	eq(buildQuery({}), "", "an empty dialog builds an empty query");
	eq(buildQuery({ words: "  ", from: " , " }), "", "and so does one with nothing but spaces");
	// what it builds must be what the parser reads
	eq(searchIndex(dated, buildQuery({ from: "a@x.com", after: "2026-06-01" })).join(","), "j3,j2", "the built query runs as itself");

	// what the mailbox is asked, versus what we filter ourselves
	const gq = (q: string) => graphSearchText(parseQuery(q));
	eq(gq("invoice"), "invoice", "words go to the mailbox");
	eq(gq('"tank management"'), '"tank management"', "and so do phrases");
	eq(gq("from:deanna subject:kore"), "from:deanna subject:kore", "from and subject are things it understands");
	eq(gq("is:unread has:attachment after:2026-01-01"), "", "the flags and the dates are not, so it is asked nothing");
	eq(gq("invoice is:unread"), "invoice", "a mixed query sends only the half it can answer");

	const pm: PCMail = { id: "x", accountId: "a", accountLabel: "l", from: "f", fromAddress: "f@x.com", subject: "s", preview: "", receivedMs: msOfKey("2026-06-15"), unread: true, flagged: true, hasAttachments: true };
	eq(passesLocalFilters(pm, parseQuery("is:unread")), true, "an unread message passes is:unread");
	eq(passesLocalFilters({ ...pm, unread: false }, parseQuery("is:unread")), false, "and a read one does not");
	eq(passesLocalFilters({ ...pm, flagged: undefined }, parseQuery("is:flagged")), false, "flagged filters server results too");
	eq(passesLocalFilters({ ...pm, hasAttachments: undefined }, parseQuery("has:attachment")), false, "so does the attachment filter");
	eq(passesLocalFilters(pm, parseQuery("after:2026-07-01")), false, "a date bound rules out what the mailbox returned anyway");
	eq(passesLocalFilters(pm, parseQuery("after:2026-01-01 before:2026-12-31")), true, "and lets through what is inside it");
	eq(passesLocalFilters(pm, parseQuery("invoice")), true, "a plain word query filters nothing here, the mailbox did that");

	// ranking and limits
	eq(searchIndex(idx, "from:irely.com", 2).join(","), "2,3", "the limit is honored, best first");
	eq(searchIndex(buildIndex([]), "anything").length, 0, "an empty index finds nothing");
	// a body that was never fetched simply cannot be searched
	const noBody = buildIndex([doc("9", "Subject only", "a@x.com", "", 1)]);
	eq(searchIndex(noBody, "subject").join(","), "9", "a message with no body is still found by its subject");
	eq(searchIndex(noBody, "unfetched").length, 0, "but a word only in its unread body is not");
}

// --- printing ---
{
	const doc = printableHtml({ subject: "Invoice T-258926", from: "Pulseway <s@p.com>", to: "Steve Palm", date: "Aug 1, 2026", bodyHtml: "<p>Thank you</p>" });
	eq(doc.startsWith("<!doctype html>"), true, "it is a whole document, not a fragment");
	eq(doc.includes("<title>Invoice T-258926</title>"), true, "the subject titles the page");
	eq(doc.includes("<p>Thank you</p>"), true, "the body goes in whole");
	eq(doc.includes("Pulseway &lt;s@p.com&gt;"), true, "angle brackets in a sender are escaped, not rendered");
	eq(doc.includes("@page"), true, "and it carries print margins of its own");

	// a subject is attacker-controlled text and this is where it becomes markup
	const nasty = printableHtml({ subject: `<img src=x onerror="alert(1)">`, from: "a", to: "b", date: "c", bodyHtml: "<p>x</p>" });
	eq(nasty.includes("<img src=x"), false, "a subject cannot smuggle a tag into the page");
	eq(nasty.includes("&lt;img src=x onerror=&quot;alert(1)&quot;&gt;"), true, "it is shown as the text it is");

	// a plain-text body must not have its markup guessed at
	const plain = printableHtml({ subject: "s", from: "a", to: "b", date: "c", bodyHtml: "line 1\n<not a tag>", plain: true });
	eq(plain.includes("&lt;not a tag&gt;"), true, "plain text is escaped rather than parsed");
	eq(plain.includes("<pre class=\"plain\">"), true, "and kept in a block that respects its line breaks");

	// a missing field simply does not get a row
	const spare = printableHtml({ subject: "s", from: "a", to: "", date: "", bodyHtml: "" });
	eq(spare.includes("<th>To</th>"), false, "an empty field is left out rather than printed blank");
	eq(spare.includes("<th>From</th>"), true, "and the ones that have a value stay");
}

// --- printing a calendar ---
{
	const days = [
		{ heading: "Monday, Aug 3", events: [{ when: "9:00 AM", title: "Standup", where: "Teams" }] },
		{ heading: "Tuesday, Aug 4", events: [] },
	];
	const a = printableAgendaHtml("Week of Aug 3", days);
	eq(a.includes("<title>Week of Aug 3</title>"), true, "the range titles the page");
	eq(a.includes("Standup"), true, "events are listed");
	eq(a.includes("Teams"), true, "with where they are");
	eq(a.includes("Nothing scheduled"), true, "an empty day says so rather than vanishing");
	eq(a.includes("page-break-inside: avoid"), true, "and a day does not split across pages");
	const noWhere = printableAgendaHtml("t", [{ heading: "d", events: [{ when: "9", title: "x" }] }]);
	eq(noWhere.includes("class=\"where\""), false, "an event with no location gets no empty line for one");

	const weeks = [
		[
			{ label: "27", dim: true, events: [] },
			{ label: "28", dim: false, events: [{ when: "9:00", title: "Review" }] },
		],
	];
	const mth = printableMonthHtml("August 2026", ["Mon", "Tue"], weeks);
	eq(mth.includes("<th>Mon</th>"), true, "the weekday headings are there");
	eq(mth.includes("size: landscape"), true, "a month grid asks for landscape rather than hoping");
	eq(mth.includes('class="dim"'), true, "days outside the month are marked");
	eq(mth.includes("Review"), true, "and the events are in their cells");

	// the same escaping duty everywhere text from elsewhere becomes markup
	const nasty = printableAgendaHtml("<b>t</b>", [{ heading: "<i>d</i>", events: [{ when: "9", title: "<script>x</script>", where: "<u>w</u>" }] }]);
	eq(nasty.includes("<script>") || nasty.includes("<i>d</i>") || nasty.includes("<u>w</u>"), false, "an event title, heading, or place cannot smuggle a tag in");
	eq(nasty.includes("&lt;script&gt;"), true, "they print as the text they are");
	const nastyM = printableMonthHtml("t", ["<b>M</b>"], [[{ label: "<i>1</i>", dim: false, events: [{ when: "9", title: "<img>" }] }]]);
	eq(nastyM.includes("<b>M</b>") || nastyM.includes("<i>1</i>") || nastyM.includes("<img>"), false, "and the same holds for the grid");
}

// --- printing a list ---
{
	const rows = [
		{ from: "Craig Cannon", subject: "PenTest report", date: "Jul 28", folder: "IT" },
		{ from: "Chris Tate", subject: "Nirvana Review", date: "Jul 31" },
	];
	const t = printableTableHtml("Unread", rows);
	eq(t.includes("<title>Unread</title>"), true, "the list's name titles the page");
	eq(t.includes("2 messages"), true, "the count is stated");
	eq(printableTableHtml("x", [rows[0]]).includes("1 message<"), true, "and reads singular for one");
	eq(printableTableHtml("x", []).includes("0 messages"), true, "an empty list still prints, saying so");
	eq((t.match(/<tr>/g) ?? []).length, 3, "a heading row and one per message");
	eq(t.includes("<td>Craig Cannon</td>"), true, "senders are cells");
	eq(t.includes("<td>IT</td>"), true, "and so is the folder when there is one");
	eq(t.includes("<td></td>"), true, "a message with no folder gets an empty cell rather than the word undefined");
	eq(t.includes("display: table-header-group"), true, "the headings repeat on later pages");
	// the same escaping duty as the message printer
	const nasty = printableTableHtml("<b>t</b>", [{ from: "a", subject: `<script>x</script>`, date: "d" }]);
	eq(nasty.includes("<script>"), false, "a subject cannot smuggle a tag into the table");
	eq(nasty.includes("&lt;script&gt;x&lt;/script&gt;"), true, "it prints as the text it is");
	eq(nasty.includes("<title>&lt;b&gt;t&lt;/b&gt;</title>"), true, "and the title is escaped too");
}

// --- print options ---
{
	// the page rule is the one place orientation and margins are decided
	eq(pageRule(undefined, false, 14), "@page { margin: 14mm; }", "no options prints portrait at the style's margin");
	eq(pageRule(undefined, true, 10), "@page { size: landscape; margin: 10mm; }", "a style that wants landscape gets it unasked");
	eq(pageRule({ landscape: true }, false, 14), "@page { size: landscape; margin: 14mm; }", "and asking turns a portrait style sideways");
	eq(pageRule({ landscape: false }, true, 10), "@page { margin: 10mm; }", "asking for portrait turns the month grid back");
	eq(pageRule({ fontPt: 20 }, false, 14), "@page { margin: 14mm; }", "size has nothing to do with the paper");

	// Normal is exactly the size each style always printed at
	eq(scaledPt(12, "m"), 12, "Normal leaves a message where it was");
	eq(scaledPt(8.5, "m"), 8.5, "and leaves the month grid where it was");
	eq(scaledPt(12, "xl"), 16.2, "Largest scales up, to a tenth of a point");
	eq(scaledPt(10.5, "xs"), 7.9, "Smallest scales down and still reads as a number");
	eq(scaledPt(12, "nonsense"), 12, "a scale nobody has heard of changes nothing");
	eq(
		PRINT_SCALES.map((s) => s.id).join(","),
		"xs,s,m,l,xl",
		"the steps run smallest to largest, so a picker built from them reads in order"
	);

	// every document flows both options through, and omitting them prints
	// what it printed before any of this existed
	const msg = { subject: "s", from: "a", to: "b", date: "c", bodyHtml: "<p>x</p>" };
	eq(printableHtml(msg).includes("@page { margin: 14mm; }"), true, "a message defaults to portrait");
	eq(printableHtml(msg).includes("font-size: 12pt"), true, "at 12pt");
	eq(printableHtml(msg, { fontPt: 16.2, landscape: true }).includes("font-size: 16.2pt"), true, "the chosen size reaches the message stylesheet");
	eq(printableHtml(msg, { fontPt: 16.2, landscape: true }).includes("size: landscape"), true, "and so does the orientation");

	const rows = [{ from: "a", subject: "s", date: "d" }];
	eq(printableTableHtml("t", rows).includes("font-size: 10pt"), true, "the table keeps its own base size");
	eq(printableTableHtml("t", rows, { fontPt: 8.8 }).includes("font-size: 8.8pt"), true, "and takes a new one");
	eq(printableTableHtml("t", rows, { landscape: true }).includes("size: landscape"), true, "a long subject line can have the paper sideways");

	const days = [{ heading: "d", events: [{ when: "9 AM", title: "x" }] }];
	eq(printableAgendaHtml("t", days).includes("font-size: 10.5pt"), true, "the agenda keeps its base size");
	eq(printableAgendaHtml("t", days, { fontPt: 12.1 }).includes("font-size: 12.1pt"), true, "and takes a new one");

	const weeks = [[{ label: "1", dim: false, events: [{ when: "9", title: "x" }] }]];
	eq(printableMonthHtml("t", ["Mon"], weeks).includes("size: landscape"), true, "the month grid still defaults to sideways");
	eq(printableMonthHtml("t", ["Mon"], weeks, { landscape: false }).includes("size: landscape"), false, "unless you ask for it upright");
	eq(printableMonthHtml("t", ["Mon"], weeks, { fontPt: 11.5 }).includes("font-size: 11.5pt"), true, "and it scales like the rest");
}

// --- signatures ---
{
	const sigs = [
		{ id: "s1", name: "iRely", html: "<p>Best regards</p>" },
		{ id: "s2", name: "Short", html: "<p>Steve</p>" },
	];
	const use = [{ accountId: "work", newId: "s1", replyId: "s2" }];
	eq(signatureFor(sigs, use, "work", "new")?.name, "iRely", "new mail takes the account's new signature");
	eq(signatureFor(sigs, use, "work", "reply")?.name, "Short", "and a reply takes its own, which is usually shorter");
	eq(signatureFor(sigs, use, "personal", "new"), null, "an account with no setting gets none");
	eq(signatureFor(sigs, [{ accountId: "work", newId: "", replyId: "s2" }], "work", "new"), null, "an empty id means none, deliberately");
	// a deleted signature must not silently become a different one
	eq(signatureFor(sigs, [{ accountId: "work", newId: "gone", replyId: "s2" }], "work", "new"), null, "an id naming nothing resolves to nothing, never to another signature");
	eq(signatureFor([], use, "work", "new"), null, "no signatures at all is no signature");

	// the one-time carry from the old single-signature setting
	const moved = migrateSignature("<p>old</p>", [], ["work", "personal"], "fresh")!;
	eq(moved.sigs.length, 1, "an old signature becomes one named signature");
	eq(moved.sigs[0].name, "Signature", "with a name to show in the list");
	eq(moved.sigs[0].html, "<p>old</p>", "and its markup intact");
	eq(moved.use.length, 2, "every account is pointed at it");
	eq(moved.use[0].newId === "fresh" && moved.use[0].replyId === "fresh", true, "for both new mail and replies, as it behaved before");
	eq(migrateSignature("", [], ["work"], "fresh"), null, "nothing to carry does nothing");
	eq(migrateSignature("   ", [], ["work"], "fresh"), null, "and neither does whitespace");
	eq(migrateSignature("<p>old</p>", sigs, ["work"], "fresh"), null, "an existing list is never overwritten by the old setting");

	// data-url images become inline attachments, which is the only way a
	// logo actually renders for the person receiving it
	const one = extractInlineImages('<p>Hi</p><img src="data:image/png;base64,AAAA">', "x");
	eq(one.images.length, 1, "a data url is pulled out");
	eq(one.images[0].contentType, "image/png", "keeping its type");
	eq(one.images[0].base64, "AAAA", "and its bytes");
	eq(one.images[0].name, "image1.png", "with a name the extension matches");
	eq(one.html, '<p>Hi</p><img src="cid:pd-x-1">', "and the markup now points at the attachment");
	const two = extractInlineImages(`<img src='data:image/jpeg;base64,BBBB'><img src="data:image/gif;base64,CCCC">`, "y");
	eq(two.images.map((i) => i.cid), ["pd-y-1", "pd-y-2"], "several images get distinct ids");
	eq(two.images.map((i) => i.name), ["image1.jpeg", "image2.gif"], "each named for its own type");
	eq(two.html.includes("data:"), false, "and none of the data urls survive");
	eq(extractInlineImages('<img src="https://x.com/logo.png">', "z").images.length, 0, "a hosted image is left exactly as it was");
	eq(extractInlineImages("<p>no images</p>", "z").html, "<p>no images</p>", "markup with no images is untouched");
}

// --- new mail arrivals ---
{
	const m = (id: string, ms: number, extra: Partial<PCMail> = {}): PCMail => ({
		id,
		accountId: "a1",
		accountLabel: "Work",
		from: `Sender ${id}`,
		fromAddress: `${id}@x.com`,
		subject: `Subject ${id}`,
		preview: "",
		receivedMs: ms,
		unread: true,
		...extra,
	});
	const none = new Set<string>();
	const list = [m("a", 100), m("b", 300), m("c", 200)];

	eq(newArrivals(list, 150, none, false).map((x) => x.id), ["c", "b"], "only what arrived after the baseline, oldest first");
	eq(newArrivals(list, 0, none, false).length, 3, "a zero baseline lets everything through");
	eq(newArrivals(list, 999, none, false).length, 0, "a baseline past everything announces nothing");
	// the guard that stops a first sync shouting about a thousand old messages
	eq(newArrivals(list, Math.max(...list.map((x) => x.receivedMs)), none, false).length, 0, "baselining on the newest silences the whole first load");
	eq(newArrivals(list, 150, new Set(["b"]), false).map((x) => x.id), ["c"], "anything already announced is not announced twice");
	eq(newArrivals([m("d", 400, { unread: false })], 0, none, false).length, 0, "mail already read is not news");
	eq(newArrivals([m("e", 400, { focused: false })], 0, none, true).length, 0, "with focused-only on, Outlook's Other is quiet");
	eq(newArrivals([m("f", 400, { focused: false })], 0, none, false).length, 1, "and with it off, everything speaks");
	eq(newArrivals([m("g", 400)], 0, none, true).length, 1, "no verdict at all still counts as focused, never silently dropped");

	eq(arrivalSummary([m("a", 1)]).title, "Sender a", "one arrival leads with who it is from");
	eq(arrivalSummary([m("a", 1)]).body, "Subject a", "and says what it is about");
	eq(arrivalSummary([m("a", 1), m("b", 2)]).title, "2 new messages", "several are counted");
	eq(arrivalSummary([m("a", 1), m("b", 2)]).body, "Sender a, Sender b", "and named");
	const many = [m("a", 1), m("b", 2), m("c", 3), m("d", 4)];
	eq(arrivalSummary(many).body, "Sender a, Sender b and 2 others", "a crowd is summarized rather than listed");
	const dupes = [m("a", 1), { ...m("b", 2), from: "Sender a" }];
	eq(arrivalSummary(dupes).body, "Sender a", "two from one person name them once");
}

// --- the split inbox ---
{
	const m = (fromAddress: string, extra: Partial<PCMail> = {}): PCMail => ({
		id: fromAddress,
		accountId: "a1",
		accountLabel: "Work",
		from: fromAddress.split("@")[0],
		fromAddress,
		subject: "s",
		preview: "",
		receivedMs: 1,
		unread: false,
		...extra,
	});

	// machines
	eq(automatedSender("no-reply@pulseway.com"), true, "no-reply is a machine");
	eq(automatedSender("noreply@x.com"), true, "so is noreply without the dash");
	eq(automatedSender("DoNotReply@x.com"), true, "and do-not-reply, whatever the case");
	eq(automatedSender("jira@irely.atlassian.net"), true, "ticketing counts");
	eq(automatedSender("notifications@github.com"), true, "so do notifications");
	eq(automatedSender("builds@ci.x.com"), true, "and build mail");
	eq(automatedSender("alerts+prod@x.com"), true, "a plus tag does not hide the name");
	// people at those same places are still people
	eq(automatedSender("steve.palm@irely.com"), false, "a person is not a machine");
	eq(automatedSender("ardi@irely.atlassian.net"), false, "a colleague mailing from atlassian is still a colleague");
	eq(automatedSender("noreplacement@x.com"), false, "a name that merely starts with 'no' is not no-reply");
	eq(automatedSender("systematic@x.com"), false, "nor is 'systematic' the system account");
	eq(automatedSender(""), false, "an empty address is not a machine");
	eq(automatedSender("bare-name"), false, "and neither is something with no domain at all");

	// sections
	eq(sectionOf(m("a@x.com", { flagged: true })), "priority", "a flag outranks everything");
	eq(sectionOf(m("no-reply@x.com", { flagged: true })), "priority", "even on a machine's mail");
	eq(sectionOf(m("a@x.com", { priority: true })), "priority", "so does high importance");
	eq(sectionOf(m("no-reply@x.com")), "notifications", "machines land in notifications");
	eq(sectionOf(m("a@x.com", { focused: true })), "focused", "Outlook's focused verdict is honored");
	eq(sectionOf(m("a@x.com", { focused: false })), "other", "and so is its other verdict");
	eq(sectionOf(m("a@x.com")), "focused", "no verdict at all counts as focused, never hidden");
	eq(sectionOf(m("no-reply@x.com", { focused: false })), "notifications", "a machine is notifications even when Outlook says other");

	// the split itself
	const list = [m("no-reply@x.com"), m("b@x.com", { focused: false }), m("c@x.com", { flagged: true }), m("d@x.com", { focused: true }), m("e@x.com", { focused: true })];
	const secs = splitSections(list);
	eq(secs.map((s) => s.key).join(","), "priority,focused,notifications,other", "sections come in reading order");
	eq(secs.find((s) => s.key === "focused")!.messages.length, 2, "a section holds everything that matched it");
	eq(
		secs.find((s) => s.key === "focused")!.messages.map((x) => x.id).join(","),
		"d@x.com,e@x.com",
		"and keeps the order the list handed it"
	);
	eq(splitSections([m("a@x.com", { focused: true })]).length, 1, "empty sections are dropped");
	eq(splitSections([]).length, 0, "an empty list splits into nothing");
	eq(
		splitSections(list).reduce((n, s) => n + s.messages.length, 0),
		list.length,
		"every message lands in exactly one section"
	);
}

// --- recipient autocomplete ---
{
	const seen = [
		{ name: "Steve Palm", email: "steve.palm@irely.com", ms: 300 },
		{ name: "", email: "STEVE.PALM@irely.com", ms: 500 },
		{ name: "Steve P", email: "steve.palm@irely.com", ms: 100 },
		{ name: "Deanna Palm", email: "deanna@irely.com", ms: 400 },
		{ name: "Webmaster", email: "webmaster@stevecorp.com", ms: 200 },
		{ name: "Nope", email: "not-an-address", ms: 900 },
	];
	const idx = rankContacts(seen);
	eq(idx.length, 3, "an address with no @ is not a contact");
	eq(idx[0].email, "steve.palm@irely.com", "the most-seen address leads");
	eq(idx[0].count, 3, "sightings of one address fold into one entry");
	eq(idx[0].email === idx[0].email.toLowerCase(), true, "addresses normalize to lower case");
	eq(idx[0].name, "Steve Palm", "a later sighting with no name leaves the name it had");
	eq(idx[0].lastMs, 500, "but the recency still moves");

	const two = rankContacts([
		{ name: "A", email: "a@x.com", ms: 1 },
		{ name: "B", email: "b@x.com", ms: 900 },
	]);
	eq(two[0].email, "b@x.com", "an equal count falls back to recency, newest first");
	eq(rankContacts([{ name: "Late", email: "c@x.com", ms: 5 }, { name: "", email: "c@x.com", ms: 1 }])[0].name, "Late", "the newest sighting that has a name wins");
	eq(rankContacts([{ name: "", email: "d@x.com", ms: 5 }, { name: "Old", email: "d@x.com", ms: 1 }])[0].name, "Old", "and a nameless newer sighting still takes an older name");

	// matching
	eq(matchContacts(idx, "ste")[0].email, "steve.palm@irely.com", "a prefix beats a match in the middle");
	eq(matchContacts(idx, "ste").length, 2, "but the middle match is still offered");
	eq(matchContacts(idx, "palm").map((c) => c.email).includes("deanna@irely.com"), true, "a surname finds the person");
	eq(matchContacts(idx, "deanna@")[0].email, "deanna@irely.com", "a partial address matches");
	eq(matchContacts(idx, "zzz").length, 0, "nothing matching gives nothing");
	eq(matchContacts(idx, "").length, 3, "an empty query offers the whole ranked list");
	eq(matchContacts(idx, "irely", 2).length, 2, "the limit is honored");

	// the address book merged with the correspondence
	{
		const seen2 = [
			{ name: "Deanna Palm", email: "deanna@irely.com", count: 40, lastMs: 900 },
			{ name: "amazon", email: "orders@amazon.com", count: 5, lastMs: 800 },
		];
		const saved = [
			{ name: "Deanna J Palm", email: "DEANNA@irely.com", company: "iRely", title: "Analyst", phone: "555-1234" },
			{ name: "Never Mailed", email: "new@x.com", company: "Acme" },
			{ name: "Broken", email: "not-an-address" },
		];
		const merged = mergePeople(seen2, saved);
		eq(merged.length, 3, "an address with no @ is not a person");
		eq(merged[0].email, "deanna@irely.com", "the one you deal with most leads");
		eq(merged[0].saved, true, "and is marked as saved once the address book confirms her");
		eq(merged[0].name, "Deanna J Palm", "the address book's name wins over one scraped off a From line");
		eq(merged[0].count, 40, "while the real correspondence count is kept");
		eq(merged[0].company, "iRely", "with the details only the address book has");
		eq(merged[0].phone, "555-1234", "including the phone");
		eq(merged.find((p) => p.email === "new@x.com")?.saved, true, "a contact never written to still appears");
		eq(merged.find((p) => p.email === "new@x.com")?.count, 0, "with no correspondence to its name");
		eq(merged[merged.length - 1].email, "new@x.com", "and sits after everyone actually corresponded with");
		eq(merged.find((p) => p.email === "orders@amazon.com")?.saved, false, "someone only ever seen in mail is not marked saved");
		eq(mergePeople([], []).length, 0, "nothing merges to nothing");
		eq(mergePeople(seen2, []).length, 2, "no address book leaves the correspondence alone");
		eq(mergePeople([], saved).map((p) => p.name).join(","), "Broken,Deanna J Palm,Never Mailed".split(",").filter((n) => n !== "Broken").join(","), "an address book alone sorts by name");
	}

	// the fragment being typed
	eq(currentAddressFragment("bob@x.com, ste", 14).text, "ste", "the fragment is what follows the last comma");
	eq(currentAddressFragment("bob@x.com, ste", 14).start, 10, "and it starts after that comma");
	eq(currentAddressFragment("ste", 3).text, "ste", "the first address is a fragment too");
	eq(currentAddressFragment("", 0).text, "", "an empty box has an empty fragment");
	eq(currentAddressFragment("a@x.com; be", 11).text, "be", "semicolons separate as well as commas");

	// taking a suggestion
	const one = applyAddressChoice("ste", 3, "steve.palm@irely.com");
	eq(one.value, "steve.palm@irely.com, ", "a choice replaces the fragment and leaves a comma to carry on");
	eq(one.caret, one.value.length, "with the caret at the end");
	const second = applyAddressChoice("bob@x.com, ste", 14, "steve.palm@irely.com");
	eq(second.value, "bob@x.com, steve.palm@irely.com, ", "an earlier address survives untouched");
	const middle = applyAddressChoice("bob@x.com, ste, later@x.com", 14, "steve.palm@irely.com");
	eq(middle.value, "bob@x.com, steve.palm@irely.com, later@x.com", "and so does one typed after it");
	eq(middle.caret, "bob@x.com, steve.palm@irely.com, ".length, "the caret lands ready for the next name");
}

// --- the "later" presets, for snooze and schedule send ---
{
	// 2026-08-05 is a Wednesday; these are all local times
	const wedMorning = new Date(2026, 7, 5, 9, 0).getTime();
	const p = whenPresets(wedMorning);
	const by = (label: string) => p.find((x) => x.label === label);

	eq(!!by("Later today") && !!by("This evening"), true, "morning offers both of today's presets");
	eq(by("Later today")!.ms, wedMorning + 3 * 3600000, "later today is three hours out");
	eq(by("This evening")!.ms, new Date(2026, 7, 5, 18, 0).getTime(), "this evening is six in the evening");
	eq(by("Tomorrow morning")!.ms, new Date(2026, 7, 6, 8, 0).getTime(), "tomorrow morning is eight the next day");
	eq(by("Tomorrow afternoon")!.ms, new Date(2026, 7, 6, 13, 0).getTime(), "tomorrow afternoon is one the next day");
	eq(by("This weekend")!.ms, new Date(2026, 7, 8, 8, 0).getTime(), "the weekend is the coming Saturday");
	eq(by("Next week")!.ms, new Date(2026, 7, 10, 8, 0).getTime(), "next week is the coming Monday");

	// nothing in the past is ever offered
	const wedNight = new Date(2026, 7, 5, 22, 30).getTime();
	const late = whenPresets(wedNight);
	eq(!!late.find((x) => x.label === "Later today"), false, "late at night drops later today");
	eq(!!late.find((x) => x.label === "This evening"), false, "and drops this evening too");
	eq(late.every((x) => x.ms > wedNight), true, "every preset offered is still in the future");
	eq(whenPresets(wedMorning).every((x) => x.ms > wedMorning), true, "and that holds in the morning as well");

	// the weekend and the week roll over rather than landing on today
	const sat = new Date(2026, 7, 8, 9, 0).getTime();
	eq(whenPresets(sat).find((x) => x.label === "This weekend")!.ms, new Date(2026, 7, 15, 8, 0).getTime(), "asking on a Saturday means the next Saturday");
	const mon = new Date(2026, 7, 10, 9, 0).getTime();
	eq(whenPresets(mon).find((x) => x.label === "Next week")!.ms, new Date(2026, 7, 17, 8, 0).getTime(), "asking on a Monday means the next Monday");
	const sun = new Date(2026, 7, 9, 9, 0).getTime();
	eq(whenPresets(sun).find((x) => x.label === "Next week")!.ms, new Date(2026, 7, 10, 8, 0).getTime(), "asking on a Sunday means tomorrow");
	eq(whenPresets(sun).find((x) => x.label === "This weekend")!.ms, new Date(2026, 7, 15, 8, 0).getTime(), "and the weekend from a Sunday is the coming Saturday");
}

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
	// a description from the rich editor has to declare itself, or the tags
	// are shown to everyone invited instead of being rendered
	const rich = graphEventBody({ title: "t", startMs: 0, endMs: 0, allDay: false, description: "<p>Agenda</p><ul><li>One</li></ul>" }, "UTC");
	eq(rich.body, { contentType: "html", content: "<p>Agenda</p><ul><li>One</li></ul>" }, "markup goes as html");
	eq((graphEventBody({ title: "t", startMs: 0, endMs: 0, allDay: false, description: "3 < 5 and 6 > 2" }, "UTC").body as { contentType: string }).contentType, "text", "a stray angle bracket does not make it html");
	eq((graphEventBody({ title: "t", startMs: 0, endMs: 0, allDay: false, description: "" }, "UTC").body as { contentType: string }).contentType, "text", "an empty description stays text");
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
	// a hand-made order leads; everything untouched keeps its alphabetical place
	const custom = orderFolderTree(folders, "in", undefined, ["von", "darwin"]);
	eq(custom.map((t) => t.folder.name), ["Inbox", "Von", "Darwin", "Upwork", "Good", "Drafts", "Stray"], "dragged folders lead their siblings in the dragged order");
	eq(custom.map((t) => t.depth), [0, 1, 1, 1, 2, 0, 0], "and reordering does not disturb the nesting");
	const one = orderFolderTree(folders, "in", undefined, ["upwork"]);
	eq(one.map((t) => t.folder.name), ["Inbox", "Upwork", "Good", "Darwin", "Von", "Drafts", "Stray"], "ordering one folder leaves the rest alphabetical behind it");
	eq(orderFolderTree(folders, "in", undefined, []).map((t) => t.folder.name), tree.map((t) => t.folder.name), "an empty custom order changes nothing");
	eq(orderFolderTree(folders, "in", undefined, ["ghost"]).map((t) => t.folder.name), tree.map((t) => t.folder.name), "an id that is no longer a folder is ignored");
	// the inbox is pinned to the top until it is itself dragged
	eq(orderFolderTree(folders, "in", undefined, ["dr"])[0].folder.name, "Inbox", "ordering a sibling does not unseat the inbox");
	eq(orderFolderTree(folders, "in", undefined, ["dr", "in"])[0].folder.name, "Drafts", "but dragging the inbox itself does");

	// folders the mailbox runs on are not ours to rename or delete
	eq(isSystemFolder("Inbox", "in", "in"), true, "the inbox is protected by its id");
	eq(isSystemFolder("Anything", "in", "in"), true, "whatever it happens to be called");
	eq(isSystemFolder("Sent Items", "x", "in"), true, "and the rest by name");
	eq(isSystemFolder("deleted items", "x", "in"), true, "case does not matter");
	eq(isSystemFolder("  Drafts  ", "x", "in"), true, "nor does stray space");
	eq(isSystemFolder("Darwin", "darwin", "in"), false, "a folder you made is yours");
	eq(isSystemFolder("Inbox Archive", "x", "in"), false, "a name that merely contains one is not one");
	eq(isSystemFolder("", "x", "in"), false, "an empty name is not a system folder");
	eq(isSystemFolder("Inbox", "x", null), true, "the name still guards when the inbox id is unknown");

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
