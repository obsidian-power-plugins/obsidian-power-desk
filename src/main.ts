import { App, FuzzySuggestModal, ItemView, MarkdownRenderChild, Menu, Modal, Notice, Platform, Plugin, PluginSettingTab, Scope, Setting, TFile, TFolder, WorkspaceLeaf, base64ToArrayBuffer, normalizePath, requestUrl, sanitizeHTMLToDom, setIcon } from "obsidian";
import { chunk, GRAPH_BATCH_MAX,
	DayCell,
	EventDraft,
	PCEvent,
	ViewMode,
	addDays,
	clipToDay,
	daySpanKeys,
	dateKeyOf,
	dayDiff,
	dayOfWeek,
	DAYS_SHORT,
	decodeJwtPayload,
	dedupePeople,
	dragTimes,
	dueReminders,
	eventDaySpan,
	eventsOnDay,
	fmtClock,
	fmtDayHeading,
	fmtDayShort,
	fmtEventRange,
	fmtFreeSlots,
	fmtMailTime,
	fmtTimeOfMs,
	fmtZoneClock,
	freeSlotsForDay,
	GraphMailLike,
	graphMailToPC,
	graphFolderToPC,
	GraphFolderLike,
	MailFolder,
	meetingProvider,
	folderSubtreeIds,
	mergeDeltaMessages,
	padFetchWindow,
	orderFolderTree,
	PCMail,
	stripHtml,
	subjectToEventTitle,
	googleEventBody,
	googleTimesBody,
	googleToPC,
	GoogleEventLike,
	graphEventBody,
	graphSetupHint,
	graphTimesBody,
	graphToPC,
	GraphEventLike,
	groupByDay,
	isoWeekNum,
	isSpanEvent,
	keyOfDate,
	keyOfMs,
	mergeForSave,
	minutesOfMs,
	monthGrid,
	msOfKey,
	packColumns,
	PALETTE,
	paletteColor,
	parseAgendaBlock,
	parseAttendeeInput,
	periodLabel,
	renderNoteName,
	searchFolderQuery,
	avatarColor,
	avatarInitials,
	GeoHit,
	pickGeoHit,
	splitPlaceQuery,
	stripMeetingBoilerplate,
	weatherGlyph,
	RepeatKind,
	sanitizeName,
	snapMin,
	sortEvents,
	spansForRow,
	stepAnchor,
	slotConflict,
	timedOnDay,
	vaultDateSpan,
	BusyRun,
	fmtVaultDate,
	parseAvailabilityView,
	viewWindow,
	weekDays,
	workWeekDays,
} from "./core";
import { parseIcsEvents } from "./ics";
import {
	DeviceCode,
	GRAPH_READ_SCOPE,
	GRAPH_SCOPE,
	GraphError,
	MailAttachment,
	archiveMessage,
	createEvent,
	deleteEvent,
	DraftMessage,
	createDraftReply,
	deleteDraft,
	deleteMessage,
	deltaFolderMessages,
	forwardMessage,
	permanentDeleteMessage,
	replyAllMessage,
	sendDraft,
	updateDraft,
	fetchCalendarView,
	fetchFolderMessages,
	fetchFolderMessagesDeep,
	fetchMailFolders,
	getMessagesBatch,
	fetchUnreadMessages,
	getEvent,
	getInboxId,
	getMailAttachmentBytes,
	getMessage,
	listMailAttachments,
	getSchedule,
	listCalendars,
	markMessageRead,
	pollToken,
	refreshTokens,
	replyMessage,
	searchMessages,
	respondEvent,
	sendGraphMail,
	startDeviceCode,
	updateEvent,
} from "./graph";
import { GoogleError, deleteGoogleEvent, fetchGoogleEvents, getGoogleEvent, insertGoogleEvent, listGoogleCalendars, loopbackAuth, patchGoogleEvent, refreshGoogleTokens } from "./google";
import { CaldavError, discoverCalendars, fetchCollectionIcs } from "./caldav";
import { DavCollection } from "./caldavxml";

/** One calendar row in the sidebar rail, whatever source it came from. */
interface SidebarCal {
	name: string;
	color: string;
	enabled: boolean;
	setEnabled: (v: boolean) => void;
	setColor: (v: string) => void;
}

const SIDEBAR_COLOR_NAMES = ["Blue", "Red", "Green", "Orange", "Purple", "Teal", "Pink", "Olive", "Rust", "Indigo"];

const VIEW_TYPE = "power-calendar";
const VIEW_TYPE_MAIL = "power-calendar-mail";
/** The Unread search folder's virtual id in folder selections and caches. */
const UNREAD_FOLDER = "__unread__";

/* ---------------- settings ---------------- */

interface CaldavCollection {
	href: string;
	name: string;
	color: string;
	enabled: boolean;
}

interface CaldavAccount {
	id: string;
	name: string;
	serverUrl: string;
	username: string;
	password: string;
	collections: CaldavCollection[];
}

interface IcsFeed {
	id: string;
	name: string;
	url: string;
	color: string;
	enabled: boolean;
}

/** Notes with a date property, rendered as events. */
interface VaultSource {
	id: string;
	name: string;
	/** Folder prefix filter; empty means the whole vault. */
	folder: string;
	dateProp: string;
	/** Optional end date/datetime property. */
	endProp: string;
	color: string;
	enabled: boolean;
}

interface GraphCalendar {
	id: string;
	name: string;
	color: string;
	enabled: boolean;
	isDefault: boolean;
}

interface GraphAccount {
	id: string;
	/** The signed-in identity (email), read from the id_token at connect time. */
	label: string;
	/** A user-chosen friendly name; shown instead of the label wherever set. */
	nickname?: string;
	/** Optional per-account app override; empty means the shared app settings
	 *  (a work tenant's app cannot sign in a personal account, so a second
	 *  registration can ride on its own account row). */
	clientId: string;
	tenant: string;
	refresh: string;
	access: string;
	expiry: number;
	grantedScope: string;
	calendars: GraphCalendar[];
	/** Show this account's inbox in the Mail view; undefined means yes. */
	mail?: boolean;
}

interface GoogleCalendar {
	id: string;
	name: string;
	color: string;
	enabled: boolean;
	primary: boolean;
	writable: boolean;
}

interface GoogleAccount {
	id: string;
	label: string;
	/** A user-chosen friendly name; shown instead of the label wherever set. */
	nickname?: string;
	refresh: string;
	access: string;
	expiry: number;
	calendars: GoogleCalendar[];
}

interface PCSettings {
	/** The shared Azure app new Microsoft sign-ins use (per-account overrides
	 *  live on the account). Tokens live locally in data.json. */
	graphClientId: string;
	graphTenant: string;
	/** Pre-1.2 single-account fields: migrated into graphAccounts on load and
	 *  kept blank so older synced devices still find their shape. */
	graphRefresh: string;
	graphAccess: string;
	graphExpiry: number;
	graphGrantedScope: string;
	graphCalendars: GraphCalendar[];
	graphAccounts: GraphAccount[];
	/** The user's own Google Cloud OAuth client (Desktop app type). */
	googleClientId: string;
	googleClientSecret: string;
	googleAccounts: GoogleAccount[];
	caldavAccounts: CaldavAccount[];
	icsFeeds: IcsFeed[];
	vaultSources: VaultSource[];
	defaultMode: ViewMode;
	weekStartsMonday: boolean;
	use24h: boolean;
	/** Where the week/day grid scrolls to on open. */
	dayStartHour: number;
	/** How many consecutive days the Day view shows, 1-7. */
	dayViewDays: number;
	/** The calendar view's left rail: mini month plus per-calendar toggles. */
	sidebarOpen: boolean;
	/** The sidebar's calendar list folds away to give the agenda room. */
	sidebarCalsCollapsed: boolean;
	sidebarAgendaCollapsed: boolean;
	/** Open-Meteo coordinates for the sidebar agenda's weather; empty = off.
	 *  Filled by looking up weatherPlace, or by hand. */
	weatherLat: string;
	weatherLon: string;
	/** The resolved place name the coordinates came from. */
	weatherPlace: string;
	weatherUnit: "f" | "c";
	/** The header filter: an unchecked class of events hides from every view.
	 *  All true = no filter applied. */
	filterMeetings: boolean;
	filterAppointments: boolean;
	filterAllDay: boolean;
	filterOnline: boolean;
	filterTentative: boolean;
	/** Outlook categories currently hidden by the Filter menu. */
	filterHiddenCategories: string[];
	/** IANA zone drawn as a second column in the time gutter; empty = off. */
	secondTimeZone: string;
	/** The window Copy free slots reads, in hours. */
	freeFromHour: number;
	freeToHour: number;
	showDeclined: boolean;
	/** Background wash for invites the user has not accepted yet. */
	calNeedsActionColor: string;
	showWeekNumbers: boolean;
	/** Minutes before a meeting to raise a sticky notice; 0 = off. */
	reminderMinutes: number;
	/** 0 = manual refresh only. */
	refreshMinutes: number;
	/** How many days back the mail sync pulls. Also sets how far Power
	 *  Assistant's Ask-your-email window can reach, since it indexes only what
	 *  is cached here. */
	mailHistoryDays: number;
	/** How many messages to retain per folder after each sync. The other half
	 *  of the Ask window's depth: a big day range still shows only this many
	 *  of the newest messages. */
	mailMaxMessages: number;
	agendaDays: number;
	notesFolder: string;
	/** Where "Save to note" writes an email. Empty falls back to the calendar
	 *  notes folder, so existing setups are unchanged; point it at a protected
	 *  folder to keep saved mail encrypted on Dropbox. */
	mailNotesFolder: string;
	noteNameTemplate: string;
	notesInNewTab: boolean;
	/** The Mail view's unread filter (Outlook's Unread search folder). */
	mailUnreadOnly: boolean;
	/** Collapsed mail folder ids, remembered like Outlook remembers them. */
	mailCollapsed: string[];
	/** When selecting a message marks it read: immediately, after a delay,
	 *  when the selection moves on, or never. */
	markRead: "select" | "delay" | "change" | "manual";
	markReadSeconds: number;
	/** While the unread filter is on, nothing auto-marks read. */
	unreadFilterKeepsUnread: boolean;
	/** Pinned folders shown first in the mail pane, in the user's own order;
	 *  name overrides the display only, never the real folder, and indent
	 *  tucks a favorite visually under the one above it. */
	mailFavorites: { accountId: string; folderId: string; name?: string; indent?: boolean }[];
	/** Saved searches rendered as Outlook-style Search Folders per account.
	 *  type and param remember the catalog choice for editing; query is what
	 *  actually runs. */
	mailSearchFolders: { accountId: string; id: string; name: string; query: string; type?: string; param?: string }[];
	/** Folders tucked out of the tree; each account lists its own under a
	 *  'Hidden folders' row for bringing back. */
	mailHiddenFolders: { accountId: string; folderId: string }[];
	/** Where 'Save to folder' writes, outside the vault; empty = Downloads. */
	mailSaveFolder: string;
	/** Scan incoming mail for orders and bills and hand matches to Power
	 *  Assistant. Off by default: it spends an AI call per matched message. */
	txnScan: boolean;
	/** Appended to new mail and replies, under a blank line. */
	mailSignature: string;
	/** Where person pages live; empty borrows Power Assistant's folder. */
	peopleFolder: string;
	/** The view phones open with (agenda reads best on a narrow screen). */
	phoneDefaultMode: ViewMode;
}

const DEFAULT_SETTINGS: PCSettings = {
	graphClientId: "",
	graphTenant: "",
	graphRefresh: "",
	graphAccess: "",
	graphExpiry: 0,
	graphGrantedScope: "",
	graphCalendars: [],
	graphAccounts: [],
	googleClientId: "",
	googleClientSecret: "",
	googleAccounts: [],
	caldavAccounts: [],
	icsFeeds: [],
	vaultSources: [],
	defaultMode: "week",
	weekStartsMonday: true,
	use24h: false,
	dayStartHour: 7,
	dayViewDays: 1,
	sidebarOpen: true,
	sidebarCalsCollapsed: false,
	sidebarAgendaCollapsed: false,
	weatherLat: "",
	weatherLon: "",
	weatherPlace: "",
	weatherUnit: "f",
	filterMeetings: true,
	filterAppointments: true,
	filterAllDay: true,
	filterOnline: true,
	filterTentative: true,
	filterHiddenCategories: [],
	secondTimeZone: "",
	freeFromHour: 9,
	freeToHour: 17,
	showDeclined: false,
	calNeedsActionColor: "#f2c94c",
	showWeekNumbers: false,
	reminderMinutes: 5,
	refreshMinutes: 5,
	mailHistoryDays: 45,
	mailMaxMessages: 50,
	agendaDays: 30,
	notesFolder: "Calendar",
	mailNotesFolder: "",
	noteNameTemplate: "{{date}} {{title}}",
	notesInNewTab: false,
	mailUnreadOnly: false,
	mailCollapsed: [],
	markRead: "select",
	markReadSeconds: 5,
	unreadFilterKeepsUnread: true,
	mailFavorites: [],
	mailSearchFolders: [],
	mailHiddenFolders: [],
	mailSaveFolder: "",
	txnScan: false,
	mailSignature: "",
	peopleFolder: "",
	phoneDefaultMode: "agenda",
};

const freshId = () => Math.random().toString(36).slice(2, 10);

/* ---------------- sources ---------------- */

type SourceDef =
	| { kind: "m365"; key: string; label: string; account: GraphAccount; calendarId: string | null; writable: boolean; color: string }
	| { kind: "google"; key: string; label: string; account: GoogleAccount; calendarId: string; writable: boolean; color: string }
	| { kind: "caldav"; key: string; label: string; account: CaldavAccount; coll: CaldavCollection; color: string }
	| { kind: "ics"; key: string; label: string; feed: IcsFeed; color: string }
	| { kind: "vault"; key: string; label: string; source: VaultSource; color: string };

interface SourceState {
	events: PCEvent[];
	error: string | null;
	fetchedAt: number;
	fromMs: number;
	toMs: number;
	inFlight: boolean;
}

/** The header shape Power Assistant's transaction API expects. Declared here
 *  rather than imported: the two plugins ship as separate bundles, so the
 *  contract is structural and feature-detected at runtime. */
interface TxnMailLike {
	id: string;
	from: string;
	subject: string;
	date: string;
	hasAttachments?: boolean;
}

export default class PowerDeskPlugin extends Plugin {
	settings: PCSettings = DEFAULT_SETTINGS;
	refreshSettingsTab: (() => void) | null = null;
	/** Per-source fetched events, keyed by SourceDef.key. */
	private cache = new Map<string, SourceState>();
	/** Open views re-render when a fetch lands. */
	listeners = new Set<() => void>();
	/** One sign-in at a time, across the wizard and settings reconnects. */
	graphConnecting = false;
	private googleConnecting = false;
	private saveTimer: number | null = null;
	/** A write of our settings is in flight. Pairs with saveTimer to cover the
	 *  whole change-to-disk span; see persistSettings(). */
	private saving = false;
	/** The settings as they last stood on disk, read or written by us. Whatever
	 *  differs from this in memory is OUR change, and only those keys may
	 *  overwrite a synced data.json; see persistSettings(). */
	private baseline: PCSettings = DEFAULT_SETTINGS;
	private externalReloadTimer: number | null = null;
	private autoTimer: number | null = null;

	async onload() {
		this.adoptSettings(Object.assign({}, DEFAULT_SETTINGS, await this.loadData()));
		// baseline is the DISK state, cloned before any migration, so the
		// migrated keys read as our change and actually persist
		this.baseline = structuredClone(this.settings);
		// pre-1.2 single-account connections migrate into the accounts list;
		// the legacy keys stay in the file, blank, for older synced devices
		if (this.settings.graphRefresh && !this.settings.graphAccounts.length) {
			this.settings.graphAccounts = [
				{
					id: freshId(),
					label: "Microsoft 365",
					clientId: "",
					tenant: "",
					refresh: this.settings.graphRefresh,
					access: this.settings.graphAccess,
					expiry: this.settings.graphExpiry,
					grantedScope: this.settings.graphGrantedScope || GRAPH_READ_SCOPE,
					calendars: this.settings.graphCalendars,
				},
			];
			this.settings.graphRefresh = "";
			this.settings.graphAccess = "";
			this.settings.graphExpiry = 0;
			this.settings.graphGrantedScope = "";
			this.settings.graphCalendars = [];
			this.queueSave();
		}
		await this.loadCacheFile();
		const stale = this.settings.graphAccounts.filter((a) => a.refresh && a.grantedScope !== GRAPH_SCOPE);
		if (stale.length) {
			new Notice(`Power Desk: reconnect ${stale.map((a) => this.nameOf(a)).join(", ")} in settings to enable the newest permissions (event editing, mail, reply windows).`, 10000);
		}

		this.registerView(VIEW_TYPE, (leaf) => new PowerCalendarView(leaf, this));
		this.registerView(VIEW_TYPE_MAIL, (leaf) => new MailView(leaf, this));
		this.addRibbonIcon("calendar-days", "Open Power Desk", () => void this.openCalendarView());
		this.addRibbonIcon("mail", "Open Power Desk inbox", () => void this.openMailView());
		this.addCommand({ id: "open-mail", icon: "inbox", name: "Open inbox", callback: () => void this.openMailView() });
		this.addCommand({
			id: "scan-transactions", icon: "receipt",
			name: "Scan mail for orders and bills now",
			callback: () => {
				if (!this.assistantTxn()) {
					new Notice("Power Desk: install Power Assistant and set a transaction folder to capture orders.", 8000);
					return;
				}
				new Notice("Power Desk: scanning mail for orders and bills...");
				void this.scanForTransactions(true).then((n) => {
					if (!n) new Notice("Power Desk: no new orders or bills found.");
				});
			},
		});

		this.addCommand({ id: "open", icon: "calendar", name: "Open calendar", callback: () => void this.openCalendarView() });
		this.addCommand({
			id: "today", icon: "calendar-days",
			name: "Go to today",
			callback: () => void this.openCalendarView().then((v) => v?.goToday()),
		});
		this.addCommand({ id: "refresh", icon: "refresh-cw", name: "Refresh calendars", callback: () => this.refreshAllOpen(true) });
		this.addCommand({ id: "connect-m365", icon: "user-plus", name: "Add a Microsoft 365 account", callback: () => new GraphAccountWizard(this.app, this).open() });
		this.addCommand({ id: "connect-google", icon: "user-plus", name: "Add a Google account", callback: () => void this.connectGoogle() });
		this.addCommand({ id: "find-event", icon: "search", name: "Find event", callback: () => new EventFindModal(this.app, this).open() });
		this.addCommand({ id: "copy-free-slots", icon: "copy", name: "Copy free slots (next 5 workdays)", callback: () => void this.copyFreeSlots() });
		this.addCommand({ id: "create-events-base", icon: "database", name: "Create events base", callback: () => void this.createEventsBase() });
		this.addCommand({ id: "test-reminder", name: "Show a test meeting reminder", callback: () => this.previewReminder() });

		// ```power-calendar``` (and its Power Desk alias) renders that day's
		// agenda inside any note; with no date line it reads the day out of the
		// note's own filename, so one template block serves every daily note.
		// The old name stays registered forever: it lives in note templates.
		for (const lang of ["power-calendar", "power-desk"]) {
			this.registerMarkdownCodeBlockProcessor(lang, (source, el, ctx) => {
				ctx.addChild(new AgendaBlock(el, this, parseAgendaBlock(source), ctx.sourcePath));
			});
		}

		this.addSettingTab(new PCSettingTab(this.app, this));
		this.watchDataFile();
		this.scheduleAutoRefresh();
		this.register(() => {
			if (this.autoTimer != null) window.clearInterval(this.autoTimer);
		});
		this.registerInterval(window.setInterval(() => this.checkReminders(), 30_000));
		this.app.workspace.onLayoutReady(() => this.checkReminders());
		this.registerEvent(this.app.metadataCache.on("changed", () => this.queueVaultRefresh()));
		this.registerEvent(this.app.vault.on("rename", () => this.queueVaultRefresh()));
		this.registerEvent(this.app.vault.on("delete", () => this.queueVaultRefresh()));
		this.register(() => {
			if (this.vaultRefreshTimer != null) window.clearTimeout(this.vaultRefreshTimer);
		});
	}

	/* ---------------- meeting reminders ---------------- */

	private remindersFired = new Set<string>();
	private remindersDay = "";

	private checkReminders() {
		const s = this.settings;
		if (s.reminderMinutes <= 0 || !this.sources().length) return;
		const today = keyOfDate(new Date());
		if (today !== this.remindersDay) {
			this.remindersDay = today;
			this.remindersFired.clear();
		}
		// open views keep the cache warm; with none open, keep today fetched
		if (!this.app.workspace.getLeavesOfType(VIEW_TYPE).length) this.ensureWindow(today, today, false);
		const due = dueReminders(this.eventsForWindow(today, today), Date.now(), s.reminderMinutes, this.remindersFired);
		for (const ev of due) {
			this.remindersFired.add(ev.id);
			this.showReminder(ev);
		}
	}

	/** The reminder card on demand. It otherwise appears minutes before a
	 *  meeting and then not again until the next one, which is no way to judge
	 *  whether it reads well, or to check that a lead time you just changed is
	 *  the one you wanted. Borrows the next real meeting where there is one, so
	 *  the preview carries the titles and join links this vault actually
	 *  produces rather than a tidy invention, and moves its start to the lead
	 *  time so the card says exactly what it will say when it fires for real.
	 *  The preview id is its own, so previewing never spends the real one. */
	private previewReminder() {
		const lead = Math.max(1, this.settings.reminderMinutes || 5);
		const today = keyOfDate(new Date());
		// with no view open today may not be cached; warm it the way the real
		// check does rather than make the preview wait on a fetch
		if (!this.app.workspace.getLeavesOfType(VIEW_TYPE).length) this.ensureWindow(today, today, false);
		const next = this.eventsForWindow(today, today)
			.filter((ev) => !ev.allDay && !ev.declined && ev.startMs > Date.now())
			.sort((a, b) => a.startMs - b.startMs)[0];
		const startMs = Date.now() + lead * 60_000;
		this.showReminder(
			next
				? { ...next, id: `preview:${next.id}`, startMs, endMs: startMs + (next.endMs - next.startMs) }
				: {
						// nothing on the calendar today: a stand-in, whose Join opens
						// Zoom's own check-your-camera page rather than a made-up
						// meeting id that can only fail
						id: "preview:sample",
						sourceId: "preview",
						title: "Sample meeting (test reminder)",
						allDay: false,
						startMs,
						endMs: startMs + 30 * 60_000,
						joinUrl: "https://zoom.us/test",
					}
		);
	}

	/** Snooze timers, held so unloading the plugin cannot leave one to fire a
	 *  reminder into a dead instance. */
	private snoozeTimers = new Set<number>();

	/** Put the reminder back later. Five minutes, or the start itself when that
	 *  comes first: a reminder that returns after the meeting began is an alarm
	 *  clock going off at noon. The floor keeps a snooze taken with seconds to
	 *  go from bouncing straight back. */
	private snoozeReminder(ev: PCEvent) {
		const delay = Math.max(30_000, Math.min(5 * 60_000, ev.startMs - Date.now()));
		const id = window.setTimeout(() => {
			this.snoozeTimers.delete(id);
			this.showReminder(ev);
		}, delay);
		this.snoozeTimers.add(id);
		const mins = Math.round(delay / 60_000);
		new Notice(`Reminder back in ${mins < 1 ? "under a minute" : `${mins} minute${mins === 1 ? "" : "s"}`}.`, 3000);
	}

	/** A sticky notice, the way a meeting reminder should be: it waits to be
	 *  dismissed, and joining is one click. Laid out as a card rather than as a
	 *  paragraph, because the three things it has to answer — which meeting, how
	 *  long have I got, how do I get in — are answered at three different
	 *  glances, and a run-on line makes all three take the same effort. */
	private showReminder(ev: PCEvent) {
		const mins = Math.max(0, Math.round((ev.startMs - Date.now()) / 60000));
		const now = mins <= 0;
		const provider = meetingProvider(ev.joinUrl);
		// the join link is the location on most invites; printing it as well as
		// putting it behind the button is the whole reason the old line sprawled
		const where = ev.location && ev.location !== ev.joinUrl && !/^https?:\/\//i.test(ev.location) ? ev.location : provider;
		let notice: Notice | null = null;
		const frag = createFragment((f) => {
			const card = f.createDiv({ cls: "pcal-remind" });
			const head = card.createDiv({ cls: "pcal-remind-head" });
			setIcon(head.createDiv({ cls: "pcal-remind-icon" }), ev.joinUrl ? "video" : "calendar-clock");
			head.createDiv({ cls: "pcal-remind-title", text: ev.title });
			const meta = card.createDiv({ cls: "pcal-remind-sub" });
			const chip = meta.createSpan({ cls: "pcal-remind-when", text: now ? "Starting now" : `In ${mins} minute${mins === 1 ? "" : "s"}` });
			chip.toggleClass("is-now", now);
			meta.createSpan({ cls: "pcal-remind-at", text: fmtTimeOfMs(ev.startMs, this.settings.use24h) });
			if (where) meta.createSpan({ cls: "pcal-remind-where", text: where });
			const row = card.createDiv({ cls: "pcal-remind-btns" });
			if (ev.joinUrl) {
				row.createEl("button", { text: "Join", cls: "mod-cta" }).addEventListener("click", () => {
					window.open(ev.joinUrl, "_blank");
					notice?.hide(); // the reminder has done its job; leaving it up is litter
				});
			}
			row.createEl("button", { cls: "pcal-remind-later", text: "Snooze" }).addEventListener("click", () => {
				notice?.hide();
				this.snoozeReminder(ev);
			});
			row.createEl("button", { cls: "pcal-remind-later", text: "Dismiss" }).addEventListener("click", () => notice?.hide());
		});
		notice = new Notice(frag, 0);
		// styling the notice itself, not just its contents: the default black
		// slab is what makes a reminder look like an error message. pw-self-styled
		// is the family's marker for a notice that brings its own surface, so the
		// light-theme restyling below leaves it alone.
		notice.noticeEl.addClass("pcal-remind-notice", "pw-self-styled");
	}

	onunload() {
		for (const id of this.snoozeTimers) window.clearTimeout(id);
		this.snoozeTimers.clear();
		if (this.saveTimer != null) {
			window.clearTimeout(this.saveTimer);
			void this.persistSettings();
		}
		if (this.cachePersistTimer != null) {
			window.clearTimeout(this.cachePersistTimer);
			void this.persistCacheFile();
		}
	}

	/* ---------------- settings persistence (synced data.json holds tokens,
	 * so a stale device writing the whole object back would log others out;
	 * every save re-reads disk and carries only our changed keys) ---------------- */

	queueSave() {
		if (this.saveTimer != null) window.clearTimeout(this.saveTimer);
		this.saveTimer = window.setTimeout(() => {
			this.saveTimer = null;
			void this.persistSettings();
		}, 400);
	}

	/**
	 * The one write path for settings. Every save goes through here so the
	 * data.json watcher can tell our own write from someone else's edit.
	 * `saving` is set synchronously before the first await, so it and
	 * `saveTimer` together cover the whole change-to-disk span with no hole.
	 */
	/**
	 * Take on new settings CONTENTS without swapping the object.
	 *
	 * Settings tabs and modals capture this object once (`const s =
	 * plugin.settings`, then `s.key = v`), so replacing it strands every one of
	 * those writes on an orphan and the setting silently stops sticking. Every
	 * assignment to this.settings goes through here for that reason. The field
	 * starts life as DEFAULT_SETTINGS itself, which must never be mutated.
	 */
	private adoptSettings(next: PCSettings) {
		if (this.settings && this.settings !== DEFAULT_SETTINGS) Object.assign(this.settings, next);
		else this.settings = { ...next };
	}

	async persistSettings() {
		this.saving = true;
		try {
			const disk = (await this.loadData()) as Partial<PCSettings> | null;
			this.adoptSettings(mergeForSave(this.settings, this.baseline, disk));
			await this.saveData(this.settings);
			this.baseline = structuredClone(this.settings);
		} finally {
			this.saving = false;
		}
	}

	private busySaving(): boolean {
		return this.saveTimer != null || this.saving;
	}

	/** Obsidian's signal that data.json changed underneath us (Sync landing
	 *  another device's write). The only such signal mobile gets. */
	async onExternalSettingsChange() {
		await this.adoptExternalData();
	}

	/** Desktop: watch the plugin folder so external edits are adopted quickly. */
	private watchDataFile() {
		if (!Platform.isDesktopApp) return;
		const basePath = (this.app.vault.adapter as unknown as { basePath?: string }).basePath;
		if (!basePath) return;
		try {
			const fs = require("node:fs") as typeof import("node:fs");
			const dir = [basePath, this.app.vault.configDir, "plugins", this.manifest.id].join("/");
			const watcher = fs.watch(dir, (_evt, name) => {
				if (name && name.toString() !== "data.json") return;
				if (this.externalReloadTimer != null) window.clearTimeout(this.externalReloadTimer);
				this.externalReloadTimer = window.setTimeout(() => {
					this.externalReloadTimer = null;
					void this.adoptExternalData();
				}, 300);
			});
			this.register(() => watcher.close());
		} catch {
			/* watcher unavailable; onExternalSettingsChange still covers sync */
		}
	}

	private async adoptExternalData() {
		if (this.busySaving()) return; // a live in-app change is on its way to disk; it wins
		const before = JSON.stringify(this.settings);
		const raw = (await this.loadData()) as Partial<PCSettings> | null;
		if (!raw) return;
		// memory can move while we await; adopting a stale read would revert it
		if (this.busySaving() || JSON.stringify(this.settings) !== before) return;
		const next = Object.assign({}, DEFAULT_SETTINGS, raw);
		if (JSON.stringify(next) === JSON.stringify(this.settings)) return; // our own write echoing back
		this.adoptSettings(next);
		this.baseline = structuredClone(next);
		this.cache.clear(); // sources or tokens may have changed shape
		this.scheduleAutoRefresh();
		this.refreshSettingsTab?.();
		this.notify();
	}

	/* ---------------- the disk cache: instant paint, refresh behind ---------------- */

	private cachePersistTimer: number | null = null;

	private cachePath(): string {
		return `${this.app.vault.configDir}/plugins/${this.manifest.id}/cache.json`;
	}

	/** Yesterday's data now, today's in a moment: the fetch caches reload at
	 *  startup so the calendar and the inbox paint instantly, while the normal
	 *  staleness rules refresh everything in the background. Tokens never land
	 *  here (data.json owns them); this file is disposable. Vault sources skip
	 *  it, since they recompute locally faster than they would deserialize. */
	private async loadCacheFile() {
		try {
			const raw = await this.app.vault.adapter.read(this.cachePath());
			const c = JSON.parse(raw) as {
				v?: number;
				events?: Record<string, { events: PCEvent[]; fetchedAt: number; fromMs: number; toMs: number }>;
				mail?: Record<string, { messages: PCMail[]; fetchedAt: number; deltaLink?: string | null }>;
				folders?: Record<string, { folders: MailFolder[]; inboxId: string | null; fetchedAt: number }>;
				folderLists?: Record<string, { messages: PCMail[]; fetchedAt: number; deltaLink?: string | null }>;
				bodies?: Record<string, { text: string; toLine: string }>;
			};
			if (c.v !== 1 && c.v !== 2) return;
			for (const [k, st] of Object.entries(c.events ?? {})) {
				if (!k.startsWith("vault:")) this.cache.set(k, { ...st, error: null, inFlight: false });
			}
			for (const [k, st] of Object.entries(c.mail ?? {})) this.mailCache.set(k, { deltaLink: null, ...st, error: null, inFlight: false });
			for (const [k, st] of Object.entries(c.folders ?? {})) this.mailFolderCache.set(k, { ...st, inFlight: false });
			for (const [k, st] of Object.entries(c.folderLists ?? {})) this.folderCache.set(k, { deltaLink: null, ...st, error: null, inFlight: false });
			for (const [k, b] of Object.entries(c.bodies ?? {})) this.bodyCache.set(k, b);
		} catch {
			/* no cache yet, or unreadable: a cold start, exactly as before */
		}
	}

	queueCachePersist() {
		if (this.cachePersistTimer != null) window.clearTimeout(this.cachePersistTimer);
		this.cachePersistTimer = window.setTimeout(() => {
			this.cachePersistTimer = null;
			void this.persistCacheFile();
		}, 8000);
	}

	private async persistCacheFile() {
		try {
			const events: Record<string, unknown> = {};
			for (const [k, st] of this.cache) if (!k.startsWith("vault:")) events[k] = { events: st.events, fetchedAt: st.fetchedAt, fromMs: st.fromMs, toMs: st.toMs };
			const mail: Record<string, unknown> = {};
			for (const [k, st] of this.mailCache) mail[k] = { messages: st.messages, fetchedAt: st.fetchedAt, deltaLink: st.deltaLink ?? null };
			const folders: Record<string, unknown> = {};
			for (const [k, st] of this.mailFolderCache) folders[k] = { folders: st.folders, inboxId: st.inboxId, fetchedAt: st.fetchedAt };
			const folderLists: Record<string, unknown> = {};
			for (const [k, st] of this.folderCache) folderLists[k] = { messages: st.messages, fetchedAt: st.fetchedAt, deltaLink: st.deltaLink ?? null };
			// the newest bodies ride along, so read messages open instantly next launch
			const bodies: Record<string, unknown> = {};
			for (const [k, b] of [...this.bodyCache].slice(-100)) bodies[k] = b;
			await this.app.vault.adapter.write(this.cachePath(), JSON.stringify({ v: 2, events, mail, folders, folderLists, bodies }));
		} catch {
			/* a failed cache write costs a cold start, nothing more */
		}
	}

	/* ---------------- source compilation and the sync engine ---------------- */

	/** Every configured, enabled source, in a stable order. Palette colors are
	 *  assigned by position so an uncolored source keeps its hue between
	 *  renders. */
	sources(): SourceDef[] {
		const out: SourceDef[] = [];
		let i = 0;
		for (const a of this.settings.graphAccounts) {
			if (!a.refresh) continue;
			const w = this.canWriteAccount(a);
			const cals = a.calendars.filter((c) => c.enabled);
			if (cals.length) {
				for (const c of cals) {
					out.push({ kind: "m365", key: `m365:${a.id}:${c.id}`, label: c.name, account: a, calendarId: c.id, writable: w, color: c.color || paletteColor(i) });
					i++;
				}
			} else {
				out.push({ kind: "m365", key: `m365:${a.id}:default`, label: this.nameOf(a), account: a, calendarId: null, writable: w, color: paletteColor(i) });
				i++;
			}
		}
		for (const g of this.settings.googleAccounts) {
			if (!g.refresh) continue;
			for (const c of g.calendars) {
				if (!c.enabled) continue;
				out.push({ kind: "google", key: `g:${g.id}:${c.id}`, label: c.name, account: g, calendarId: c.id, writable: c.writable, color: c.color || paletteColor(i) });
				i++;
			}
		}
		for (const account of this.settings.caldavAccounts) {
			for (const coll of account.collections) {
				if (!coll.enabled) continue;
				out.push({ kind: "caldav", key: `dav:${account.id}:${coll.href}`, label: coll.name, account, coll, color: coll.color || paletteColor(i) });
				i++;
			}
		}
		for (const feed of this.settings.icsFeeds) {
			if (!feed.enabled) continue;
			out.push({ kind: "ics", key: `ics:${feed.id}`, label: feed.name, feed, color: feed.color || paletteColor(i) });
			i++;
		}
		for (const v of this.settings.vaultSources) {
			if (!v.enabled) continue;
			out.push({ kind: "vault", key: `vault:${v.id}`, label: v.name, source: v, color: v.color || paletteColor(i) });
			i++;
		}
		// three accounts can each ship a calendar literally named "Calendar";
		// a duplicated label borrows its account's name so chips stay readable
		const counts = new Map<string, number>();
		for (const d of out) counts.set(d.label, (counts.get(d.label) ?? 0) + 1);
		for (const d of out) {
			if ((counts.get(d.label) ?? 0) < 2) continue;
			const owner = d.kind === "m365" || d.kind === "google" ? this.nameOf(d.account) : d.kind === "caldav" ? d.account.name : "";
			if (owner) d.label = `${d.label} · ${owner}`;
		}
		return out;
	}

	/* ---------------- weather (Open-Meteo, keyless) ---------------- */

	private weatherCache: { fetchedAt: number; days: Map<string, { hi: number; lo: number; code: number }> } | null = null;
	private weatherInFlight = false;

	/** The forecast for one day key, from a cache refreshed hourly. Null while
	 *  weather is off or the first fetch is still out. */
	weatherFor(key: string): { hi: number; lo: number; code: number } | null {
		const s = this.settings;
		if (!s.weatherLat.trim() || !s.weatherLon.trim()) return null;
		if (!this.weatherCache || Date.now() - this.weatherCache.fetchedAt > 3_600_000) void this.fetchWeather();
		return this.weatherCache?.days.get(key) ?? null;
	}

	clearWeather() {
		this.weatherCache = null;
	}

	/** A city name to coordinates via Open-Meteo's keyless geocoder. The comma
	 *  form ('Flower Mound, TX') searches on the city and picks the candidate
	 *  matching the region, since the geocoder rejects the whole string. */
	async geocodePlace(query: string): Promise<{ lat: string; lon: string; label: string } | null> {
		const { name, region } = splitPlaceQuery(query);
		if (!name) return null;
		try {
			const r = await requestUrl({ url: `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(name)}&count=10&language=en&format=json` });
			const hit = pickGeoHit(region, (r.json as { results?: GeoHit[] }).results ?? []);
			if (!hit) return null;
			return { lat: String(hit.latitude), lon: String(hit.longitude), label: [hit.name, hit.admin1, hit.country_code].filter(Boolean).join(", ") };
		} catch {
			return null;
		}
	}

	private async fetchWeather() {
		if (this.weatherInFlight) return;
		const lat = this.settings.weatherLat.trim();
		const lon = this.settings.weatherLon.trim();
		if (!lat || !lon) return;
		this.weatherInFlight = true;
		try {
			const unit = this.settings.weatherUnit === "c" ? "celsius" : "fahrenheit";
			const r = await requestUrl({
				url: `https://api.open-meteo.com/v1/forecast?latitude=${encodeURIComponent(lat)}&longitude=${encodeURIComponent(lon)}&daily=weather_code,temperature_2m_max,temperature_2m_min&timezone=auto&forecast_days=10&temperature_unit=${unit}`,
			});
			const d = r.json as { daily?: { time?: string[]; weather_code?: number[]; temperature_2m_max?: number[]; temperature_2m_min?: number[] } };
			const days = new Map<string, { hi: number; lo: number; code: number }>();
			(d.daily?.time ?? []).forEach((t, i) => {
				days.set(t, { hi: Math.round(d.daily?.temperature_2m_max?.[i] ?? 0), lo: Math.round(d.daily?.temperature_2m_min?.[i] ?? 0), code: d.daily?.weather_code?.[i] ?? 0 });
			});
			this.weatherCache = { fetchedAt: Date.now(), days };
			this.notify();
		} catch {
			// a failed fetch keeps what it had and retries in an hour
			this.weatherCache = { fetchedAt: Date.now(), days: this.weatherCache?.days ?? new Map() };
		} finally {
			this.weatherInFlight = false;
		}
	}

	/** Sources changed shape (added, removed, toggled): drop stale caches and
	 *  let every open view refetch what it looks at. */
	sourcesChanged() {
		const live = new Set(this.sources().map((d) => d.key));
		for (const key of Array.from(this.cache.keys())) if (!live.has(key)) this.cache.delete(key);
		this.notify();
	}

	notify() {
		for (const l of Array.from(this.listeners)) l();
	}

	/** Make sure every source has fresh events covering the window; fetches run
	 *  in the background and notify views as each lands. */
	ensureWindow(fromKey: string, toKey: string, force = false) {
		const fromMs = msOfKey(fromKey);
		const toMs = msOfKey(addDays(toKey, 1));
		const staleMs = Math.max(1, this.settings.refreshMinutes) * 60000;
		for (const def of this.sources()) {
			const st = this.cache.get(def.key);
			if (st?.inFlight) continue;
			const covers = st && st.fromMs <= fromMs && st.toMs >= toMs;
			const fresh = st && Date.now() - st.fetchedAt < staleMs;
			if (!force && covers && fresh) continue;
			if (!force && covers && st?.error) continue; // don't hammer a broken server on every render
			// a covered cache refreshes over its own stored span, so the wide
			// window never shrinks back to the visible one; a miss fetches the
			// window plus padding, so stepping week to week stays off the network
			const span = covers && st ? { fromMs: st.fromMs, toMs: st.toMs } : padFetchWindow(fromMs, toMs);
			void this.fetchSource(def, span.fromMs, span.toMs);
		}
	}

	anyInFlight(): boolean {
		for (const st of this.cache.values()) if (st.inFlight) return true;
		return false;
	}

	sourceStates(): { def: SourceDef; st: SourceState | null }[] {
		return this.sources().map((def) => ({ def, st: this.cache.get(def.key) ?? null }));
	}

	private async fetchSource(def: SourceDef, fromMs: number, toMs: number) {
		const st: SourceState = this.cache.get(def.key) ?? { events: [], error: null, fetchedAt: 0, fromMs, toMs, inFlight: false };
		st.inFlight = true;
		this.cache.set(def.key, st);
		this.notify();
		try {
			let events: PCEvent[];
			if (def.kind === "m365") {
				const token = await this.graphTokenFor(def.account);
				const raw = await this.mailboxGate(def.account.id, () => fetchCalendarView(token, def.calendarId, new Date(fromMs).toISOString(), new Date(toMs).toISOString()));
				events = raw
					.map((r) => graphToPC(r as GraphEventLike, { sourceId: def.key, calendarName: def.label, color: def.color, writable: def.writable }))
					.filter((e): e is PCEvent => e != null);
			} else if (def.kind === "google") {
				const token = await this.googleTokenFor(def.account);
				const raw = await fetchGoogleEvents(token, def.calendarId, new Date(fromMs).toISOString(), new Date(toMs).toISOString());
				events = raw
					.map((r) => googleToPC(r as GoogleEventLike, { sourceId: def.key, calendarName: def.label, color: def.color, writable: def.writable }))
					.filter((e): e is PCEvent => e != null);
			} else if (def.kind === "caldav") {
				const payloads = await fetchCollectionIcs(def.coll.href, def.account.username, def.account.password, fromMs, toMs);
				events = [];
				for (const p of payloads) {
					try {
						events.push(...parseIcsEvents(p, fromMs, toMs, { sourceId: def.key, calendarName: def.label, color: def.color }));
					} catch {
						/* one unreadable resource must not hide the rest */
					}
				}
			} else if (def.kind === "ics") {
				const url = def.feed.url.trim().replace(/^webcal:\/\//i, "https://");
				const r = await requestUrl({ url, throw: false });
				if (r.status >= 400) throw new Error(`The feed answered ${r.status}.`);
				events = parseIcsEvents(r.text, fromMs, toMs, { sourceId: def.key, calendarName: def.label, color: def.color });
			} else {
				events = this.vaultEvents(def, fromMs, toMs);
			}
			st.events = events;
			st.error = null;
			st.fromMs = fromMs;
			st.toMs = toMs;
		} catch (e) {
			// keep whatever rendered before; stale beats blank while a server sulks
			st.error = e instanceof Error ? e.message : String(e);
		} finally {
			st.fetchedAt = Date.now();
			st.inFlight = false;
			this.notify();
			this.queueCachePersist();
		}
	}

	/** The dated notes of one vault source as events. Synchronous reads over
	 *  the metadata cache, so "fetching" is instant and free. */
	private vaultEvents(def: Extract<SourceDef, { kind: "vault" }>, fromMs: number, toMs: number): PCEvent[] {
		const out: PCEvent[] = [];
		const folder = def.source.folder.trim() ? normalizePath(def.source.folder.trim()) : "";
		for (const f of this.app.vault.getMarkdownFiles()) {
			if (folder && f.path !== folder && !f.path.startsWith(folder + "/")) continue;
			const fm = this.app.metadataCache.getFileCache(f)?.frontmatter as Record<string, unknown> | undefined;
			const raw = fm?.[def.source.dateProp];
			if (raw == null) continue;
			const span = vaultDateSpan(raw, def.source.endProp ? fm?.[def.source.endProp] : undefined);
			if (!span || span.endMs <= fromMs || span.startMs >= toMs) continue;
			out.push({
				id: `note:${f.path}`,
				sourceId: def.key,
				calendarName: def.label,
				color: def.color,
				title: f.basename,
				allDay: span.allDay,
				startMs: span.startMs,
				endMs: span.endMs,
				canEdit: !span.allDay, // a timed note can be dragged; banners stay put
				notePath: f.path,
			});
		}
		return out;
	}

	/** Vault edits land constantly; a short debounce drops the vault sources'
	 *  caches so the next render recomputes them (instant and local). */
	private vaultRefreshTimer: number | null = null;
	private queueVaultRefresh() {
		if (!this.settings.vaultSources.some((v) => v.enabled)) return;
		if (this.vaultRefreshTimer != null) window.clearTimeout(this.vaultRefreshTimer);
		this.vaultRefreshTimer = window.setTimeout(() => {
			this.vaultRefreshTimer = null;
			for (const def of this.sources()) if (def.kind === "vault") this.cache.delete(def.key);
			this.notify();
		}, 500);
	}

	/** Everything fetched that intersects the window, declined filtered. */
	eventsForWindow(fromKey: string, toKey: string): PCEvent[] {
		const fromMs = msOfKey(fromKey);
		const toMs = msOfKey(addDays(toKey, 1));
		const out: PCEvent[] = [];
		for (const def of this.sources()) {
			const st = this.cache.get(def.key);
			if (!st) continue;
			for (const ev of st.events) {
				if (ev.endMs <= fromMs || ev.startMs >= toMs) continue;
				if (ev.declined && !this.settings.showDeclined) continue;
				out.push(ev);
			}
		}
		return out;
	}

	refreshAllOpen(force: boolean) {
		for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE)) {
			if (leaf.view instanceof PowerCalendarView) leaf.view.refresh(force);
		}
	}

	private scheduleAutoRefresh() {
		if (this.autoTimer != null) window.clearInterval(this.autoTimer);
		this.autoTimer = null;
		if (this.settings.refreshMinutes > 0) {
			this.autoTimer = window.setInterval(() => {
				this.refreshAllOpen(true);
				if (this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL).length) this.refreshMailAll(true);
				// deliberately not gated on a view: capture has to keep working
				// while the user is anywhere else in the vault
				void this.scanForTransactions();
			}, this.settings.refreshMinutes * 60000);
		}
	}

	refreshCadenceChanged() {
		this.scheduleAutoRefresh();
	}

	/* ---------------- Microsoft 365 ---------------- */

	/** Own settings first, Power Assistant's app registration as the fallback,
	 *  so a vault already running the suite needs no second Azure setup while
	 *  this plugin still stands entirely on its own. */
	private siblingGraph(): { clientId: string; tenant: string } {
		const plugs = (this.app as unknown as { plugins?: { plugins?: Record<string, { settings?: { graphClientId?: string; graphTenant?: string } }> } }).plugins
			?.plugins;
		const s = (plugs?.["powerassistant"] ?? plugs?.["powercapture"])?.settings;
		return { clientId: s?.graphClientId?.trim() ?? "", tenant: s?.graphTenant?.trim() ?? "" };
	}

	effectiveClientId(): string {
		return this.settings.graphClientId.trim() || this.siblingGraph().clientId;
	}

	effectiveTenant(): string {
		return this.settings.graphTenant.trim() || (this.settings.graphClientId.trim() ? "common" : this.siblingGraph().tenant) || "common";
	}

	usingSiblingApp(): boolean {
		return !this.settings.graphClientId.trim() && !!this.siblingGraph().clientId;
	}

	/** The app a given account signs in through: its own override, else the
	 *  shared app settings (with the Power Assistant fallback). */
	clientIdFor(a: GraphAccount): string {
		return a.clientId.trim() || this.effectiveClientId();
	}

	tenantFor(a: GraphAccount): string {
		if (a.tenant.trim()) return a.tenant.trim();
		if (a.clientId.trim()) return "common";
		return this.effectiveTenant();
	}

	graphConnected(): boolean {
		return this.settings.graphAccounts.some((a) => !!a.refresh);
	}

	/** Whether one account can write events. Substring checks, not equality:
	 *  every scope widening must leave older grants their old powers. */
	canWriteAccount(a: GraphAccount): boolean {
		return !!a.refresh && a.grantedScope.includes("Calendars.ReadWrite");
	}

	canMailAccount(a: GraphAccount): boolean {
		return !!a.refresh && (a.grantedScope.includes("Mail.Read") || a.grantedScope.includes("Mail.ReadWrite"));
	}

	/** The user's friendly name for an account when set, else its signed-in identity. */
	nameOf(a: { nickname?: string; label: string }): string {
		return a.nickname?.trim() || a.label;
	}

	/** Anywhere to create events at all; gates the create affordances. */
	anyWritable(): boolean {
		return this.writableTargets().length > 0;
	}

	private async graphTokenFor(a: GraphAccount): Promise<string> {
		if (a.access && Date.now() < a.expiry - 60_000) return a.access;
		if (!a.refresh) throw new Error(`Reconnect ${this.nameOf(a)} in settings first.`);
		let t;
		try {
			t = await refreshTokens(this.clientIdFor(a), this.tenantFor(a), a.refresh, a.grantedScope || GRAPH_SCOPE);
		} catch (e) {
			// a rejected refresh token (invalid_grant, ...) is dead: clear it so the
			// UI reads "signed out" instead of failing every refresh silently. A
			// codeless error is a transient network blip, so keep the token.
			if (e instanceof GraphError && e.code) {
				a.access = "";
				a.refresh = "";
				a.expiry = 0;
				await this.persistSettings();
				this.refreshSettingsTab?.();
			}
			throw e;
		}
		a.access = t.access_token;
		if (t.refresh_token) a.refresh = t.refresh_token;
		a.expiry = Date.now() + t.expires_in * 1000;
		// pin the app that just proved it owns this refresh token, so a later
		// change to the shared app settings cannot orphan the account (tenant
		// first: its resolution depends on the row's clientId still being empty)
		if (!a.clientId.trim()) {
			a.tenant = this.tenantFor(a);
			a.clientId = this.clientIdFor(a);
		}
		await this.persistSettings();
		return a.access;
	}

	/** Device-code sign-in: a new account, a reconnect, or a scope upgrade.
	 *  Never handles the password. */
	async connectGraph(account?: GraphAccount, appOverride?: { clientId: string; tenant: string }) {
		const target: GraphAccount = account ?? {
			id: freshId(),
			label: "Microsoft 365",
			clientId: appOverride?.clientId.trim() ?? "",
			tenant: appOverride?.tenant.trim() ?? "",
			refresh: "",
			access: "",
			expiry: 0,
			grantedScope: "",
			calendars: [],
		};
		if (!this.clientIdFor(target)) {
			new Notice("Power Desk: enter an Azure app (client) ID in settings first.");
			return;
		}
		if (this.graphConnecting) {
			new Notice("Power Desk: a sign-in is already in progress.");
			return;
		}
		this.graphConnecting = true;
		try {
			await this.runDeviceCodeFlow(target, !account);
		} finally {
			this.graphConnecting = false;
		}
	}

	/** One device-code session for `a`: `onCode` presents the fresh code,
	 *  `cancelled` is consulted between polls. Resolves to the connected
	 *  account (persisted, calendars synced) once tokens are in, or null on
	 *  cancel or timeout; definitive sign-in failures throw for the caller's
	 *  UI. An identity that is already connected refreshes its existing row
	 *  instead of growing a duplicate, since sign-in pages love re-picking the
	 *  account that is already in the browser. */
	async deviceCodeSession(a: GraphAccount, isNew: boolean, onCode: (dc: DeviceCode) => void, cancelled: () => boolean): Promise<GraphAccount | null> {
		const dc = await startDeviceCode(this.clientIdFor(a), this.tenantFor(a));
		onCode(dc);
		const deadline = Date.now() + dc.expires_in * 1000;
		const interval = Math.max(3, dc.interval) * 1000;
		while (Date.now() < deadline && !cancelled()) {
			await sleep(interval);
			if (cancelled()) return null;
			const res = await pollToken(this.clientIdFor(a), this.tenantFor(a), dc.device_code);
			if (res === "pending") continue;
			const claims = decodeJwtPayload(res.id_token ?? "");
			const who = (claims?.preferred_username ?? claims?.email ?? claims?.name) as string | undefined;
			if (who) a.label = who;
			const target = isNew ? (this.settings.graphAccounts.find((x) => x.label.toLowerCase() === a.label.toLowerCase()) ?? a) : a;
			if (target !== a) {
				// this session's tokens belong to the app the wizard chose
				target.clientId = a.clientId;
				target.tenant = a.tenant;
			}
			target.access = res.access_token;
			target.refresh = res.refresh_token;
			target.expiry = Date.now() + res.expires_in * 1000;
			target.grantedScope = GRAPH_SCOPE;
			// pin the app this sign-in went through to the row (tenant first: its
			// resolution depends on the row's clientId still being empty), so the
			// account survives any later change to the shared app settings
			if (!target.clientId.trim()) {
				target.tenant = this.tenantFor(target);
				target.clientId = this.clientIdFor(target);
			}
			if (isNew && target === a) this.settings.graphAccounts = [...this.settings.graphAccounts, a];
			await this.persistSettings();
			await this.syncGraphCalendars(target);
			this.refreshSettingsTab?.();
			this.sourcesChanged();
			return target;
		}
		return null;
	}

	private async runDeviceCodeFlow(a: GraphAccount, isNew: boolean) {
		let modal: DeviceCodeModal | undefined;
		try {
			const acct = await this.deviceCodeSession(
				a,
				isNew,
				(dc) => {
					modal = new DeviceCodeModal(this.app, dc);
					modal.open();
				},
				() => !!modal && !modal.waiting
			);
			if (acct) new Notice(`Power Desk: connected ${this.nameOf(acct)}.`);
			else if (modal?.waiting) new Notice("Power Desk: sign-in timed out; try again.");
		} catch (e) {
			this.graphErrorNotice(e);
		}
		modal?.close();
	}

	/** Pull one account's calendar list, keeping the user's enabled flags and
	 *  color overrides for calendars already known. */
	async syncGraphCalendars(a: GraphAccount) {
		try {
			const list = await listCalendars(await this.graphTokenFor(a));
			const known = new Map(a.calendars.map((c) => [c.id, c]));
			a.calendars = list.map((c, i) => ({
				id: c.id,
				name: c.name,
				color: known.get(c.id)?.color || (/^#[0-9a-fA-F]{6}$/.test(c.hexColor) ? c.hexColor.toLowerCase() : "") || paletteColor(i),
				enabled: known.get(c.id)?.enabled ?? true,
				isDefault: c.isDefaultCalendar,
			}));
			await this.persistSettings();
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	private graphErrorNotice(e: unknown) {
		const msg = e instanceof Error ? e.message : String(e);
		const hint = graphSetupHint(msg);
		new Notice("Power Desk: " + (hint ? `${hint}\n\n(${msg})` : msg), hint ? 15000 : 8000);
	}

	removeGraphAccount(a: GraphAccount) {
		this.settings.graphAccounts = this.settings.graphAccounts.filter((x) => x.id !== a.id);
		void this.persistSettings();
		new Notice(`Power Desk: removed ${this.nameOf(a)}.`);
		this.refreshSettingsTab?.();
		this.sourcesChanged();
	}

	/* ---------------- Google ---------------- */

	googleReady(): boolean {
		return !!this.settings.googleClientId.trim() && !!this.settings.googleClientSecret.trim();
	}

	/** Loopback sign-in for a new Google account, or a reconnect of `existing`
	 *  (which keeps its calendar toggles). Desktop only; the tokens sync. */
	async connectGoogle(existing?: GoogleAccount) {
		if (!this.googleReady()) {
			new Notice("Power Desk: enter your Google client ID and secret in settings first. The README walks through the one-time Google Cloud setup.", 10000);
			return;
		}
		if (!Platform.isDesktopApp) {
			new Notice("Power Desk: sign in to Google once on desktop; the connection then syncs to this device.");
			return;
		}
		if (this.googleConnecting) {
			new Notice("Power Desk: a sign-in is already in progress.");
			return;
		}
		this.googleConnecting = true;
		new Notice("Power Desk: finish the Google sign-in in your browser.");
		try {
			const t = await loopbackAuth(this.settings.googleClientId.trim(), this.settings.googleClientSecret.trim(), (url) => window.open(url));
			if (!t.refresh_token) {
				new Notice("Power Desk: Google returned no refresh token. Remove the app's access at myaccount.google.com/permissions and connect again.", 12000);
				return;
			}
			const claims = decodeJwtPayload(t.id_token ?? "");
			const g: GoogleAccount = existing ?? { id: freshId(), label: "Google", refresh: "", access: "", expiry: 0, calendars: [] };
			g.label = (claims?.email as string) || g.label;
			g.refresh = t.refresh_token;
			g.access = t.access_token;
			g.expiry = Date.now() + t.expires_in * 1000;
			if (!existing) this.settings.googleAccounts = [...this.settings.googleAccounts, g];
			await this.persistSettings();
			new Notice(`Power Desk: connected ${this.nameOf(g)}.`);
			await this.syncGoogleCalendars(g);
			this.refreshSettingsTab?.();
			this.sourcesChanged();
		} catch (e) {
			this.googleErrorNotice(e);
		} finally {
			this.googleConnecting = false;
		}
	}

	private async googleTokenFor(g: GoogleAccount): Promise<string> {
		if (g.access && Date.now() < g.expiry - 60_000) return g.access;
		if (!g.refresh) throw new Error(`Reconnect ${this.nameOf(g)} in settings first.`);
		let t;
		try {
			t = await refreshGoogleTokens(this.settings.googleClientId.trim(), this.settings.googleClientSecret.trim(), g.refresh);
		} catch (e) {
			if (e instanceof GoogleError && e.code) {
				g.access = "";
				g.refresh = "";
				g.expiry = 0;
				await this.persistSettings();
				this.refreshSettingsTab?.();
			}
			throw e;
		}
		g.access = t.access_token;
		if (t.refresh_token) g.refresh = t.refresh_token;
		g.expiry = Date.now() + t.expires_in * 1000;
		await this.persistSettings();
		return g.access;
	}

	async syncGoogleCalendars(g: GoogleAccount) {
		try {
			const list = await listGoogleCalendars(await this.googleTokenFor(g));
			const known = new Map(g.calendars.map((c) => [c.id, c]));
			g.calendars = list.map((c, i) => ({
				id: c.id,
				name: c.name,
				color: known.get(c.id)?.color || c.color || paletteColor(i),
				enabled: known.get(c.id)?.enabled ?? true,
				primary: c.primary,
				writable: c.writable,
			}));
			await this.persistSettings();
		} catch (e) {
			this.googleErrorNotice(e);
		}
	}

	removeGoogleAccount(g: GoogleAccount) {
		this.settings.googleAccounts = this.settings.googleAccounts.filter((x) => x.id !== g.id);
		void this.persistSettings();
		new Notice(`Power Desk: removed ${this.nameOf(g)}.`);
		this.refreshSettingsTab?.();
		this.sourcesChanged();
	}

	private googleErrorNotice(e: unknown) {
		const msg = e instanceof Error ? e.message : String(e);
		const weekly =
			e instanceof GoogleError && e.code === "invalid_grant"
				? " If this happens every 7 days, open your Google Cloud project's OAuth consent screen and press Publish app."
				: "";
		new Notice("Power Desk: " + msg + weekly, weekly ? 15000 : 8000);
	}

	/* ---------------- event writes (Microsoft 365 only, phase two) ---------------- */

	private localTz(): string {
		return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
	}

	/** Force-refetch remote sources over their cached windows, so a write's
	 *  truth replaces the optimistic guess. Scoped to the touched source when
	 *  a key is given; the rest refresh on their own cadence. */
	private refetchRemote(onlyKey?: string) {
		for (const def of this.sources()) {
			if (def.kind !== "m365" && def.kind !== "google") continue;
			if (onlyKey && def.key !== onlyKey) continue;
			const st = this.cache.get(def.key);
			if (st && !st.inFlight) void this.fetchSource(def, st.fromMs, st.toMs);
		}
	}

	private sourceByKey(key: string): SourceDef | null {
		return this.sources().find((d) => d.key === key) ?? null;
	}

	private anyErrorNotice(kind: SourceDef["kind"], e: unknown) {
		if (kind === "google") this.googleErrorNotice(e);
		else this.graphErrorNotice(e);
	}

	/** Every calendar a new event can land in, across every account. Labels
	 *  carry the account only when more than one is signed in. */
	writableTargets(): { key: string; label: string; def: SourceDef }[] {
		const liveAccounts = this.settings.graphAccounts.filter((a) => a.refresh).length + this.settings.googleAccounts.filter((g) => g.refresh).length;
		const out: { key: string; label: string; def: SourceDef }[] = [];
		for (const def of this.sources()) {
			if ((def.kind === "m365" || def.kind === "google") && def.writable) {
				const nm = this.nameOf(def.account);
				const suffix = liveAccounts > 1 && def.label !== nm && !def.label.endsWith(`· ${nm}`) ? ` (${nm})` : "";
				out.push({ key: def.key, label: def.label + suffix, def });
			}
		}
		return out;
	}

	/** Quick-create's target: a default/primary calendar first, else the first
	 *  writable one anywhere. */
	defaultWriteTarget(): { key: string; label: string; def: SourceDef } | null {
		const targets = this.writableTargets();
		const isHome = (t: { def: SourceDef }) => {
			const d = t.def;
			if (d.kind === "m365") return d.calendarId == null || !!d.account.calendars.find((c) => c.id === d.calendarId)?.isDefault;
			if (d.kind === "google") return !!d.account.calendars.find((c) => c.id === d.calendarId)?.primary;
			return false;
		};
		return targets.find(isHome) ?? targets[0] ?? null;
	}

	/** Mutate the cached copy of an event so the grid answers instantly; the
	 *  refetch that follows replaces it with the server's truth. */
	private applyLocalTimes(ev: PCEvent, startMs: number, endMs: number) {
		const st = this.cache.get(ev.sourceId);
		const cached = st?.events.find((e) => e.id === ev.id);
		if (cached) {
			cached.startMs = startMs;
			cached.endMs = endMs;
		}
		this.notify();
	}

	async createEventAt(targetKey: string | null, draft: EventDraft): Promise<boolean> {
		const target = (targetKey ? this.writableTargets().find((t) => t.key === targetKey) : null) ?? this.defaultWriteTarget();
		if (!target) {
			new Notice("Power Desk: no writable calendar is connected.");
			return false;
		}
		const def = target.def;
		try {
			if (def.kind === "m365") await createEvent(await this.graphTokenFor(def.account), def.calendarId, graphEventBody(draft, this.localTz()));
			else if (def.kind === "google") await insertGoogleEvent(await this.googleTokenFor(def.account), def.calendarId, googleEventBody(draft, this.localTz()), (draft.attendees?.length ?? 0) > 0);
			else return false;
			new Notice("Power Desk: event created.");
			this.refetchRemote(def.key);
			return true;
		} catch (e) {
			this.anyErrorNotice(def.kind, e);
			return false;
		}
	}

	/** Drag-to-reschedule a vault note: rewrite its date property (and a
	 *  datetime end, keeping the drag's length) in frontmatter. */
	private async moveVaultEvent(def: Extract<SourceDef, { kind: "vault" }>, ev: PCEvent, startMs: number, endMs: number): Promise<void> {
		const file = ev.notePath ? this.app.vault.getAbstractFileByPath(ev.notePath) : null;
		if (!(file instanceof TFile)) return;
		this.applyLocalTimes(ev, startMs, endMs);
		try {
			await this.app.fileManager.processFrontMatter(file, (fm: Record<string, unknown>) => {
				fm[def.source.dateProp] = fmtVaultDate(startMs, !ev.allDay);
				if (def.source.endProp && fm[def.source.endProp] != null && /T\d{1,2}:\d{2}/.test(String(fm[def.source.endProp]))) {
					fm[def.source.endProp] = fmtVaultDate(endMs, true);
				}
			});
		} catch (e) {
			new Notice("Power Desk: could not update the note (" + (e instanceof Error ? e.message : String(e)) + ").");
			this.queueVaultRefresh();
		}
	}

	/** Reschedule (drag, resize): optimistic, then patched, then verified. */
	async moveEvent(ev: PCEvent, startMs: number, endMs: number): Promise<void> {
		const def = this.sourceByKey(ev.sourceId);
		if (def?.kind === "vault") {
			await this.moveVaultEvent(def, ev, startMs, endMs);
			return;
		}
		if (!def || (def.kind !== "m365" && def.kind !== "google")) return;
		const prev = { startMs: ev.startMs, endMs: ev.endMs };
		this.applyLocalTimes(ev, startMs, endMs);
		try {
			if (def.kind === "m365") await updateEvent(await this.graphTokenFor(def.account), ev.id, graphTimesBody(startMs, endMs, ev.allDay, this.localTz()));
			else await patchGoogleEvent(await this.googleTokenFor(def.account), def.calendarId, ev.id, googleTimesBody(startMs, endMs, ev.allDay, this.localTz()));
			this.refetchRemote(def.key);
		} catch (e) {
			this.applyLocalTimes(ev, prev.startMs, prev.endMs);
			this.anyErrorNotice(def.kind, e);
		}
	}

	async updateCalEvent(ev: PCEvent, draft: EventDraft): Promise<boolean> {
		const def = this.sourceByKey(ev.sourceId);
		if (!def || (def.kind !== "m365" && def.kind !== "google")) return false;
		try {
			if (def.kind === "m365") await updateEvent(await this.graphTokenFor(def.account), ev.id, graphEventBody(draft, this.localTz()));
			else await patchGoogleEvent(await this.googleTokenFor(def.account), def.calendarId, ev.id, googleEventBody(draft, this.localTz()), draft.attendees != null);
			new Notice("Power Desk: event updated.");
			this.refetchRemote(def.key);
			return true;
		} catch (e) {
			this.anyErrorNotice(def.kind, e);
			return false;
		}
	}

	/** The series master behind a recurring instance, mapped like any fetched
	 *  event so the editor can treat it normally. Patching it edits the series. */
	async loadSeriesMaster(ev: PCEvent): Promise<PCEvent | null> {
		const def = this.sourceByKey(ev.sourceId);
		if (!def || !ev.seriesId) return null;
		try {
			if (def.kind === "m365") {
				const raw = await getEvent(await this.graphTokenFor(def.account), ev.seriesId);
				return graphToPC(raw as GraphEventLike, { sourceId: def.key, calendarName: def.label, color: def.color, writable: def.writable });
			}
			if (def.kind === "google") {
				const raw = await getGoogleEvent(await this.googleTokenFor(def.account), def.calendarId, ev.seriesId);
				return googleToPC(raw as GoogleEventLike, { sourceId: def.key, calendarName: def.label, color: def.color, writable: def.writable });
			}
			return null;
		} catch (e) {
			this.anyErrorNotice(def.kind, e);
			return null;
		}
	}

	/** Delete an event, or cancel one occurrence of a series. Providers keep
	 *  deletions recoverable on their side (Deleted Items, Google's trash). */
	async deleteCalEvent(ev: PCEvent): Promise<void> {
		const def = this.sourceByKey(ev.sourceId);
		if (!def || (def.kind !== "m365" && def.kind !== "google")) return;
		const st = this.cache.get(ev.sourceId);
		if (st) {
			// deleting a series master sweeps its expanded instances too
			st.events = st.events.filter((e) => e.id !== ev.id && e.seriesId !== ev.id);
			this.notify();
		}
		try {
			if (def.kind === "m365") await deleteEvent(await this.graphTokenFor(def.account), ev.id);
			else await deleteGoogleEvent(await this.googleTokenFor(def.account), def.calendarId, ev.id);
			new Notice(ev.recurring ? "Power Desk: occurrence cancelled." : "Power Desk: event deleted.");
		} catch (e) {
			this.anyErrorNotice(def.kind, e);
		} finally {
			this.refetchRemote(def.key);
		}
	}

	/** Respond to an invite. Optimistic on the cached copy; the provider
	 *  carries the answer to the organizer. */
	async rsvpEvent(ev: PCEvent, response: "accepted" | "tentative" | "declined"): Promise<void> {
		const def = this.sourceByKey(ev.sourceId);
		if (!def || (def.kind !== "m365" && def.kind !== "google")) return;
		const cached = this.cache.get(ev.sourceId)?.events.find((e) => e.id === ev.id);
		if (cached) {
			cached.myResponse = response;
			cached.declined = response === "declined";
			cached.tentative = response === "tentative";
			this.notify();
		}
		try {
			if (def.kind === "m365") {
				const action = response === "accepted" ? "accept" : response === "tentative" ? "tentativelyAccept" : "decline";
				await respondEvent(await this.graphTokenFor(def.account), ev.id, action);
			} else {
				// Google has no response endpoint: read the event, flip the self
				// attendee, patch the whole list back
				const token = await this.googleTokenFor(def.account);
				const raw = await getGoogleEvent(token, def.calendarId, ev.id);
				const attendees = ((raw.attendees as Record<string, unknown>[]) ?? []).map((a) => (a.self ? { ...a, responseStatus: response } : a));
				await patchGoogleEvent(token, def.calendarId, ev.id, { attendees }, true);
			}
			new Notice("Power Desk: response sent.");
		} catch (e) {
			this.anyErrorNotice(def.kind, e);
		} finally {
			this.refetchRemote(def.key);
		}
	}

	/** Bounded wait for in-flight fetches, for commands that read the cache
	 *  right after ensuring a window. */
	private async waitForFetches(timeoutMs = 8000): Promise<void> {
		const start = Date.now();
		while (this.anyInFlight() && Date.now() - start < timeoutMs) await sleep(150);
	}

	/** The next five workdays' open slots as paste-able text on the clipboard:
	 *  the local, serverless cousin of a scheduling link. */
	async copyFreeSlots(): Promise<void> {
		const s = this.settings;
		const fromMin = s.freeFromHour * 60;
		const toMin = s.freeToHour * 60;
		if (toMin <= fromMin) {
			new Notice("Power Desk: the availability window in settings ends before it starts.");
			return;
		}
		const days: string[] = [];
		let k = keyOfDate(new Date());
		while (days.length < 5) {
			const dow = dayOfWeek(k);
			if (dow >= 1 && dow <= 5) days.push(k);
			k = addDays(k, 1);
		}
		this.ensureWindow(days[0], days[days.length - 1], false);
		await this.waitForFetches();
		const events = this.eventsForWindow(days[0], days[days.length - 1]);
		const text = fmtFreeSlots(
			days.map((key) => ({ key, slots: freeSlotsForDay(events, key, fromMin, toMin) })),
			fromMin,
			toMin,
			s.use24h
		);
		await navigator.clipboard.writeText(text);
		new Notice("Power Desk: free slots copied.\n\n" + text, 8000);
	}

	/** Free/busy runs per invitee over a window, via one getSchedule call.
	 *  Null on a hard failure (already noticed); per-person misses come back
	 *  as row errors so one outsider cannot blank the whole panel. */
	async checkAvailability(def: Extract<SourceDef, { kind: "m365" }>, emails: string[], startMs: number, endMs: number): Promise<{ email: string; runs: BusyRun[]; error: string | null }[] | null> {
		try {
			const infos = await getSchedule(await this.graphTokenFor(def.account), emails, startMs, endMs, 30);
			return infos.map((s) => ({ email: s.email, runs: parseAvailabilityView(s.availabilityView, startMs, 30), error: s.error }));
		} catch (e) {
			this.graphErrorNotice(e);
			return null;
		}
	}

	/* ---------------- mail: the triage inbox ---------------- */

	private mailCache = new Map<string, { messages: PCMail[]; fetchedAt: number; error: string | null; inFlight: boolean; deltaLink?: string | null }>();

	mailAccounts(): GraphAccount[] {
		return this.settings.graphAccounts.filter((a) => this.canMailAccount(a) && a.mail !== false);
	}

	mailChanged() {
		const live = new Set(this.mailAccounts().map((a) => a.id));
		for (const k of Array.from(this.mailCache.keys())) if (!live.has(k)) this.mailCache.delete(k);
		this.notify();
	}

	accountById(id: string): GraphAccount | null {
		return this.settings.graphAccounts.find((a) => a.id === id) ?? null;
	}

	/** Every account's recent inbox, newest first. */
	allMail(): PCMail[] {
		const out: PCMail[] = [];
		for (const a of this.mailAccounts()) out.push(...(this.mailCache.get(a.id)?.messages ?? []));
		return out.sort((x, y) => y.receivedMs - x.receivedMs);
	}

	anyMailInFlight(): boolean {
		for (const st of this.mailCache.values()) if (st.inFlight) return true;
		return false;
	}

	mailErrors(): string[] {
		const out: string[] = [];
		for (const a of this.mailAccounts()) {
			const err = this.mailCache.get(a.id)?.error;
			if (err) out.push(`${this.nameOf(a)}: ${err}`);
		}
		return out;
	}

	ensureMail(force = false) {
		const staleMs = Math.max(1, this.settings.refreshMinutes) * 60000;
		for (const a of this.mailAccounts()) {
			const st = this.mailCache.get(a.id);
			if (st?.inFlight) continue;
			if (!force && st && Date.now() - st.fetchedAt < staleMs) continue;
			void this.fetchMailFor(a);
		}
	}

	private async fetchMailFor(a: GraphAccount) {
		const st = this.mailCache.get(a.id) ?? { messages: [], fetchedAt: 0, error: null, inFlight: false, deltaLink: null };
		st.inFlight = true;
		this.mailCache.set(a.id, st);
		this.notify();
		try {
			await this.syncList(a, "inbox", st);
			st.error = null;
			void this.prefetchBodies(st.messages);
		} catch (e) {
			st.error = e instanceof Error ? e.message : String(e);
		} finally {
			st.fetchedAt = Date.now();
			st.inFlight = false;
			this.notify();
			this.queueCachePersist();
		}
	}

	/* ----- folder trees and per-folder mail ----- */

	private mailFolderCache = new Map<string, { folders: MailFolder[]; inboxId: string | null; fetchedAt: number; inFlight: boolean }>();
	private folderCache = new Map<string, { messages: PCMail[]; fetchedAt: number; error: string | null; inFlight: boolean; deltaLink?: string | null }>();

	/** One delta-backed sync of a folder's list into its cache entry: only
	 *  changes travel once a link exists. An initial round is bounded to the
	 *  recent window, topped up with a snapshot for quiet folders whose mail
	 *  is all older. Any delta misbehavior falls back to the plain snapshot,
	 *  so a tenant quirk costs efficiency, never the mail. */
	private async syncList(a: GraphAccount, folderId: string, st: { messages: PCMail[]; deltaLink?: string | null }): Promise<void> {
		const token = await this.graphTokenFor(a);
		const days = Math.min(365, Math.max(7, this.settings.mailHistoryDays || 45));
		const cap = Math.min(5000, Math.max(50, this.settings.mailMaxMessages || 50));
		const sinceMs = Date.now() - days * 86400000;
		const toPC = (raw: unknown[]) => raw.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label)).filter((m): m is PCMail => m != null);
		try {
			let link = st.deltaLink ?? null;
			let d = await deltaFolderMessages(token, folderId, link, sinceMs, cap);
			if (d.resync) {
				link = null;
				d = await deltaFolderMessages(token, folderId, null, sinceMs, cap);
			}
			let incoming = toPC(d.messages);
			// a quiet folder's mail can all predate the window; show its
			// newest anyway, exactly as a snapshot would
			if (!link && incoming.length < 10) {
				incoming = mergeDeltaMessages(toPC(await fetchFolderMessages(token, folderId, 25)), incoming, [], Math.max(75, cap));
			}
			st.messages = mergeDeltaMessages(link ? st.messages : [], incoming, d.removedIds, cap);
			st.deltaLink = d.deltaLink;
		} catch {
			st.messages = toPC(await fetchFolderMessages(token, folderId, Math.min(cap, 100)));
			st.deltaLink = null;
		}
	}

	ensureMailFolders(force = false) {
		const staleMs = Math.max(1, this.settings.refreshMinutes) * 60000 * 3; // trees change rarely
		for (const a of this.mailAccounts()) {
			const st = this.mailFolderCache.get(a.id);
			if (st?.inFlight) continue;
			if (!force && st && Date.now() - st.fetchedAt < staleMs) continue;
			void this.fetchFoldersFor(a);
		}
	}

	private async fetchFoldersFor(a: GraphAccount) {
		const st = this.mailFolderCache.get(a.id) ?? { folders: [], inboxId: null, fetchedAt: 0, inFlight: false };
		st.inFlight = true;
		this.mailFolderCache.set(a.id, st);
		try {
			const token = await this.graphTokenFor(a);
			const [inboxId, raw] = await Promise.all([getInboxId(token), fetchMailFolders(token)]);
			st.inboxId = inboxId;
			st.folders = raw.map((f) => graphFolderToPC(f as GraphFolderLike)).filter((f): f is MailFolder => f != null);
		} catch {
			/* folder pane just stays as it was; the inbox path already notices */
		} finally {
			st.fetchedAt = Date.now();
			st.inFlight = false;
			this.notify();
			this.queueCachePersist();
			void this.prefetchFolders(a);
		}
	}

	folderTreeFor(a: GraphAccount): { folder: MailFolder; depth: number; expandable: boolean }[] {
		const st = this.mailFolderCache.get(a.id);
		return st ? orderFolderTree(st.folders, st.inboxId, new Set(this.settings.mailCollapsed)) : [];
	}

	/** Unread across one folder and everything under it, for collapsed rows. */
	folderUnreadRollup(accountId: string, folderId: string): number {
		const st = this.mailFolderCache.get(accountId);
		if (!st) return 0;
		const ids = folderSubtreeIds(st.folders, folderId);
		let n = 0;
		for (const f of st.folders) if (ids.has(f.id)) n += f.unread;
		return n;
	}

	inboxIdFor(a: GraphAccount): string | null {
		return this.mailFolderCache.get(a.id)?.inboxId ?? null;
	}

	/* Graph allows about four concurrent requests per mailbox before it
	 * throws MailboxConcurrency throttles; this per-account gate keeps the
	 * plugin under the limit and absorbs a stray throttle with one retry. */
	private mailboxGates = new Map<string, { active: number; queue: (() => void)[] }>();

	async mailboxGate<T>(accountId: string, fn: () => Promise<T>): Promise<T> {
		const g = this.mailboxGates.get(accountId) ?? { active: 0, queue: [] };
		this.mailboxGates.set(accountId, g);
		while (g.active >= 3) await new Promise<void>((res) => g.queue.push(res));
		g.active++;
		try {
			try {
				return await fn();
			} catch (e) {
				if (e instanceof Error && /MailboxConcurrency|ApplicationThrottled|throttl/i.test(e.message)) {
					await sleep(2000);
					return await fn();
				}
				throw e;
			}
		} finally {
			g.active--;
			g.queue.shift()?.();
		}
	}

	ensureFolderMail(accountId: string, folderId: string, force = false) {
		const a = this.accountById(accountId);
		if (!a) return;
		const key = `${accountId}:${folderId}`;
		const st = this.folderCache.get(key);
		if (st?.inFlight) return;
		const staleMs = Math.max(1, this.settings.refreshMinutes) * 60000;
		if (!force && st && Date.now() - st.fetchedAt < staleMs) return;
		void this.fetchFolderMailFor(a, folderId);
	}

	private async fetchFolderMailFor(a: GraphAccount, folderId: string) {
		const key = `${a.id}:${folderId}`;
		const st = this.folderCache.get(key) ?? { messages: [], fetchedAt: 0, error: null, inFlight: false, deltaLink: null };
		st.inFlight = true;
		this.folderCache.set(key, st);
		// a cap keeps abandoned folder syncs from accumulating forever; the
		// prefetcher fills a good chunk of this on purpose
		if (this.folderCache.size > 50) {
			let oldest: string | null = null;
			let oldestAt = Infinity;
			for (const [k, v] of this.folderCache) {
				if (k !== key && !v.inFlight && v.fetchedAt < oldestAt) {
					oldest = k;
					oldestAt = v.fetchedAt;
				}
			}
			if (oldest) this.folderCache.delete(oldest);
		}
		this.notify();
		try {
			await this.mailboxGate(a.id, async () => {
				if (folderId === UNREAD_FOLDER) {
					const raw = await fetchUnreadMessages(await this.graphTokenFor(a));
					st.messages = raw.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label)).filter((m): m is PCMail => m != null);
				} else {
					await this.syncList(a, folderId, st);
				}
			});
			st.error = null;
		} catch (e) {
			st.error = e instanceof Error ? e.message : String(e);
		} finally {
			st.fetchedAt = Date.now();
			st.inFlight = false;
			this.notify();
			this.queueCachePersist();
		}
	}

	/* ----- prefetch: the Outlook cached-mode feel ----- */

	private prefetchingFolders = new Set<string>();

	/** Warm the folders you are likely to click next: the inbox subtree,
	 *  unread first, quietly and one at a time. Delta links make each later
	 *  pass nearly free. */
	private async prefetchFolders(a: GraphAccount) {
		if (this.prefetchingFolders.has(a.id)) return;
		this.prefetchingFolders.add(a.id);
		try {
			const st = this.mailFolderCache.get(a.id);
			if (!st) return;
			const subtree = folderSubtreeIds(st.folders, st.inboxId);
			const staleMs = Math.max(1, this.settings.refreshMinutes) * 60000;
			const picks = st.folders
				.filter((f) => subtree.has(f.id) && f.id !== st.inboxId && f.total > 0)
				.sort((x, y) => y.unread - x.unread || y.total - x.total)
				.slice(0, 15);
			for (const f of picks) {
				const cached = this.folderCache.get(`${a.id}:${f.id}`);
				if (cached && Date.now() - cached.fetchedAt < staleMs) continue;
				await this.fetchFolderMailFor(a, f.id);
			}
		} finally {
			this.prefetchingFolders.delete(a.id);
		}
	}

	private bodyPrefetchRunning = false;

	/** Read the top of a list before it is clicked, so the reading pane is
	 *  instant for anything visible. Silent; a miss costs nothing. */
	async prefetchBodies(list: PCMail[]) {
		if (this.bodyPrefetchRunning) return;
		this.bodyPrefetchRunning = true;
		try {
			await this.readMailBodies(list.slice(0, 40), true);
		} finally {
			this.bodyPrefetchRunning = false;
		}
	}

	/** Warm the body cache for many messages at once.
	 *
	 *  One $batch carries twenty reads, so this turns what used to be one round
	 *  trip per message into one per twenty. Everything already cached is
	 *  skipped, messages are grouped by account because each needs its own
	 *  token, and a throttled batch backs off once rather than hammering.
	 *  Returns how many bodies were newly cached. */
	async readMailBodies(list: readonly PCMail[], quiet = false): Promise<number> {
		const pending = list.filter((m) => {
			const hit = this.bodyCache.get(m.id);
			return !hit || hit.html === undefined;
		});
		if (!pending.length) return 0;

		const byAccount = new Map<string, PCMail[]>();
		for (const m of pending) {
			const arr = byAccount.get(m.accountId) ?? [];
			arr.push(m);
			byAccount.set(m.accountId, arr);
		}

		let cached = 0;
		for (const [accountId, msgs] of byAccount) {
			const a = this.accountById(accountId);
			if (!a) continue;
			let token: string;
			try {
				token = await this.graphTokenFor(a);
			} catch (e) {
				if (!quiet) this.graphErrorNotice(e);
				continue;
			}
			const byId = new Map(msgs.map((m) => [m.id, m]));
			for (const group of chunk(msgs.map((m) => m.id), GRAPH_BATCH_MAX)) {
				let out;
				try {
					out = await this.mailboxGate(accountId, () => getMessagesBatch(token, group));
				} catch (e) {
					if (!quiet) this.graphErrorNotice(e);
					continue;
				}
				for (const [id, raw] of out.ok) {
					const m = byId.get(id);
					if (!m) continue;
					this.cacheMailBody(m, raw);
					cached++;
				}
				// one polite pause on throttling; the next sync picks up the rest
				if (out.retryAfterMs > 0) {
					await sleep(Math.min(out.retryAfterMs, 10000));
					break;
				}
			}
		}
		if (cached) {
			this.queueCachePersist();
			this.notify();
		}
		return cached;
	}

	/** Turn one Graph message into a cached body entry. Shared by the single
	 *  and batched readers so both store exactly the same shape. */
	private cacheMailBody(m: PCMail, raw: Record<string, unknown>): { text: string; html?: string; toLine: string } {
		const body = raw.body as { contentType?: string; content?: string } | undefined;
		const content = body?.content ?? "";
		const isHtml = body?.contentType?.toLowerCase() === "html" || /<\w+[^>]*>/.test(content);
		const text = isHtml ? stripHtml(content) : content;
		const toLine = ((raw.toRecipients as { emailAddress?: { name?: string; address?: string } }[] | undefined) ?? [])
			.map((r) => r.emailAddress?.name || r.emailAddress?.address || "")
			.filter(Boolean)
			.join(", ");
		const result = { text, html: isHtml ? content : "", toLine };
		this.bodyCache.set(m.id, result);
		if (this.bodyCache.size > 150) {
			const oldest = this.bodyCache.keys().next().value;
			if (oldest) this.bodyCache.delete(oldest);
		}
		return result;
	}

	folderMail(accountId: string, folderId: string): PCMail[] {
		const messages = this.folderCache.get(`${accountId}:${folderId}`)?.messages ?? [];
		if (folderId !== UNREAD_FOLDER) return messages;
		// the Unread search folder scopes the mailbox-wide fetch to the inbox
		// subtree once the folder tree knows it; until then, everything shows
		const st = this.mailFolderCache.get(accountId);
		const subtree = st ? folderSubtreeIds(st.folders, st.inboxId) : new Set<string>();
		if (!subtree.size) return messages;
		return messages.filter((m) => !m.folderId || subtree.has(m.folderId));
	}

	/** Unread across the inbox and every folder under it, from tree counts. */
	unreadSubtreeCount(a: GraphAccount): number {
		const st = this.mailFolderCache.get(a.id);
		if (!st) return 0;
		const subtree = folderSubtreeIds(st.folders, st.inboxId);
		let n = 0;
		for (const f of st.folders) if (subtree.has(f.id)) n += f.unread;
		return n;
	}

	folderNamesFor(accountId: string): Map<string, string> {
		const st = this.mailFolderCache.get(accountId);
		return new Map((st?.folders ?? []).map((f) => [f.id, f.name]));
	}

	folderMailError(accountId: string, folderId: string): string | null {
		return this.folderCache.get(`${accountId}:${folderId}`)?.error ?? null;
	}

	folderMailInFlight(accountId: string, folderId: string): boolean {
		return this.folderCache.get(`${accountId}:${folderId}`)?.inFlight ?? false;
	}

	/* ----- mailbox search ----- */

	private mailSearch: { query: string; results: PCMail[]; inFlight: boolean; error: string | null } | null = null;

	mailSearchState(): { query: string; results: PCMail[]; inFlight: boolean; error: string | null } | null {
		return this.mailSearch;
	}

	/** Search every mail account's whole mailbox; one account failing costs
	 *  its results, not the search. A newer query supersedes a slower one. */
	async runMailSearch(query: string): Promise<void> {
		const q = query.trim();
		if (!q) {
			this.clearMailSearch();
			return;
		}
		const state = { query: q, results: [] as PCMail[], inFlight: true, error: null as string | null };
		this.mailSearch = state;
		this.notify();
		const all: PCMail[] = [];
		let firstError: string | null = null;
		for (const a of this.mailAccounts()) {
			try {
				const raw = await searchMessages(await this.graphTokenFor(a), q, 25);
				all.push(...raw.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label)).filter((m): m is PCMail => m != null));
			} catch (e) {
				firstError = firstError ?? (e instanceof Error ? e.message : String(e));
			}
		}
		if (this.mailSearch !== state) return;
		state.results = all.sort((x, y) => y.receivedMs - x.receivedMs);
		state.error = firstError;
		state.inFlight = false;
		this.notify();
	}

	clearMailSearch() {
		if (!this.mailSearch) return;
		this.mailSearch = null;
		this.notify();
	}

	/** Every cached list that could hold a given account's message, so read
	 *  flags and archives stay consistent between the inbox and folder views. */
	private cachedMailLists(accountId: string): { messages: PCMail[] }[] {
		const out: { messages: PCMail[] }[] = [];
		const inbox = this.mailCache.get(accountId);
		if (inbox) out.push(inbox);
		for (const [key, st] of this.folderCache) if (key.startsWith(accountId + ":")) out.push(st);
		return out;
	}

	/** The timer's one entry point: inboxes, trees, and any folder being viewed. */
	refreshMailAll(force: boolean) {
		this.ensureMail(force);
		this.ensureMailFolders(force);
		for (const key of this.folderCache.keys()) {
			const [accountId, folderId] = [key.slice(0, key.indexOf(":")), key.slice(key.indexOf(":") + 1)];
			this.ensureFolderMail(accountId, folderId, force);
		}
	}

	/** Bodies already read this session come back instantly. */
	private bodyCache = new Map<string, { text: string; html?: string; toLine: string }>();

	/** The full body for the reading pane: HTML mail keeps its HTML for the
	 *  sanitized rich rendering, with the stripped text riding along for notes
	 *  and previews. Quiet mode is for prefetch, where a failure should cost
	 *  nothing visible. */
	async readMailBody(m: PCMail, quiet = false): Promise<{ text: string; html?: string; toLine: string } | null> {
		// entries cached before HTML rendering existed lack the html field
		// entirely (even text mail now stores html: ""); refetch those once
		const hit = this.bodyCache.get(m.id);
		if (hit && hit.html !== undefined) return hit;
		const a = this.accountById(m.accountId);
		if (!a) return null;
		try {
			const raw = await this.mailboxGate(a.id, async () => getMessage(await this.graphTokenFor(a), m.id));
			const result = this.cacheMailBody(m, raw);
			this.queueCachePersist();
			return result;
		} catch (e) {
			if (!quiet) this.graphErrorNotice(e);
			return null;
		}
	}

	/** The attachments on a message, without payloads. Errors read as none. */
	async mailAttachments(m: PCMail): Promise<MailAttachment[]> {
		const a = this.accountById(m.accountId);
		if (!a) return [];
		try {
			return await this.mailboxGate(a.id, async () => listMailAttachments(await this.graphTokenFor(a), m.id));
		} catch {
			return [];
		}
	}

	/** One attachment's base64 payload; null for kinds carrying no bytes. */
	async mailAttachmentRaw(m: PCMail, attachmentId: string, quiet = false): Promise<{ name: string; contentType: string; contentBytes: string; contentId?: string } | null> {
		const a = this.accountById(m.accountId);
		if (!a) return null;
		try {
			return await this.mailboxGate(a.id, async () => getMailAttachmentBytes(await this.graphTokenFor(a), m.id, attachmentId));
		} catch (e) {
			if (!quiet) this.graphErrorNotice(e);
			return null;
		}
	}

	/** Save an attachment where the vault's attachment settings point; the
	 *  created path comes back for opening. */
	async saveMailAttachment(m: PCMail, att: MailAttachment): Promise<string | null> {
		const raw = await this.mailAttachmentRaw(m, att.id);
		if (!raw) {
			new Notice("Power Desk: this attachment carries no file to save.");
			return null;
		}
		const path = await this.app.fileManager.getAvailablePathForAttachment(raw.name);
		await this.app.vault.createBinary(path, base64ToArrayBuffer(raw.contentBytes));
		return path;
	}

	/** Save outside the vault: the configured folder, or the OS Downloads.
	 *  Desktop only, through the renderer's node require. */
	async saveMailAttachmentLocal(m: PCMail, att: MailAttachment): Promise<string | null> {
		const req = (window as unknown as { require?: (mod: string) => unknown }).require;
		if (!Platform.isDesktopApp || !req) {
			new Notice("Power Desk: saving outside the vault needs the desktop app.");
			return null;
		}
		const raw = await this.mailAttachmentRaw(m, att.id);
		if (!raw) {
			new Notice("Power Desk: this attachment carries no file to save.");
			return null;
		}
		const fs = req("fs") as { promises: { mkdir: (p: string, o: { recursive: boolean }) => Promise<unknown>; writeFile: (p: string, d: Uint8Array) => Promise<void> }; existsSync: (p: string) => boolean };
		const pathMod = req("path") as { join: (...parts: string[]) => string; extname: (p: string) => string; basename: (p: string, ext?: string) => string };
		const os = req("os") as { homedir: () => string };
		const dir = this.settings.mailSaveFolder.trim() || pathMod.join(os.homedir(), "Downloads");
		try {
			await fs.promises.mkdir(dir, { recursive: true });
			let target = pathMod.join(dir, raw.name);
			for (let n = 1; fs.existsSync(target); n++) {
				const ext = pathMod.extname(raw.name);
				target = pathMod.join(dir, `${pathMod.basename(raw.name, ext)} (${n})${ext}`);
			}
			await fs.promises.writeFile(target, new Uint8Array(base64ToArrayBuffer(raw.contentBytes)));
			new Notice(`Power Desk: saved ${target}.`);
			return target;
		} catch (e) {
			new Notice("Power Desk: could not save there. " + (e instanceof Error ? e.message : String(e)));
			return null;
		}
	}

	async setMailRead(m: PCMail, read: boolean) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			const cached = list.messages.find((x) => x.id === m.id);
			if (cached) cached.unread = !read;
		}
		this.notify();
		try {
			await markMessageRead(await this.graphTokenFor(a), m.id, read);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async archiveMail(m: PCMail) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			list.messages = list.messages.filter((x) => x.id !== m.id);
		}
		this.notify();
		try {
			await archiveMessage(await this.graphTokenFor(a), m.id);
		} catch (e) {
			this.graphErrorNotice(e);
			void this.fetchMailFor(a);
		}
	}

	/** The configured signature as HTML: raw when it already is HTML, escaped
	 *  with line breaks when plain. Empty when none is set. */
	signatureHtml(): string {
		const sig = this.settings.mailSignature.trim();
		if (!sig) return "";
		if (/<[a-z][\s\S]*>/i.test(sig)) return sig;
		const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		return `<p>${esc(sig).replace(/\n/g, "<br>")}</p>`;
	}

	/** The typed text as HTML with the signature beneath: an HTML signature
	 *  rides untouched, a plain one gets its line breaks. */
	composeHtml(text: string): string {
		const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		const body = `<p>${esc(text).replace(/\n/g, "<br>")}</p>`;
		const sigHtml = this.signatureHtml();
		return sigHtml ? `${body}<br>${sigHtml}` : body;
	}

	/* ----- the rich compose window's draft plumbing ----- */

	async createMailDraft(m: PCMail, kind: "reply" | "replyAll" | "forward"): Promise<DraftMessage | null> {
		const a = this.accountById(m.accountId);
		if (!a) return null;
		try {
			return await this.mailboxGate(a.id, async () => createDraftReply(await this.graphTokenFor(a), m.id, kind));
		} catch (e) {
			this.graphErrorNotice(e);
			return null;
		}
	}

	async sendMailDraft(accountId: string, draftId: string, patch: { subject: string; html: string; to: string[]; cc: string[] }): Promise<boolean> {
		const a = this.accountById(accountId);
		if (!a) return false;
		try {
			const token = await this.graphTokenFor(a);
			await updateDraft(token, draftId, patch);
			await sendDraft(token, draftId);
			new Notice("Power Desk: mail sent.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	async discardMailDraft(accountId: string, draftId: string) {
		const a = this.accountById(accountId);
		if (!a) return;
		try {
			await deleteDraft(await this.graphTokenFor(a), draftId);
		} catch {
			/* a stray draft in Drafts is harmless */
		}
	}

	async replyToMail(m: PCMail, text: string): Promise<boolean> {
		const a = this.accountById(m.accountId);
		if (!a) return false;
		try {
			await replyMessage(await this.graphTokenFor(a), m.id, this.composeHtml(text));
			new Notice("Power Desk: reply sent.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** Like Outlook's Delete: into Deleted Items, recoverable there. */
	async deleteMail(m: PCMail) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			list.messages = list.messages.filter((x) => x.id !== m.id);
		}
		this.notify();
		try {
			await deleteMessage(await this.graphTokenFor(a), m.id);
		} catch (e) {
			this.graphErrorNotice(e);
			void this.fetchMailFor(a);
		}
	}

	/** Shift+Del: past Deleted Items entirely, unrecoverable. */
	async permanentDeleteMail(m: PCMail) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			list.messages = list.messages.filter((x) => x.id !== m.id);
		}
		this.notify();
		try {
			await permanentDeleteMessage(await this.graphTokenFor(a), m.id);
		} catch (e) {
			this.graphErrorNotice(e);
			void this.fetchMailFor(a);
		}
	}

	async replyAllToMail(m: PCMail, text: string): Promise<boolean> {
		const a = this.accountById(m.accountId);
		if (!a) return false;
		try {
			await replyAllMessage(await this.graphTokenFor(a), m.id, this.composeHtml(text));
			new Notice("Power Desk: reply sent to everyone.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** Forward through Graph, so the original body and attachments ride along. */
	async forwardMail(m: PCMail, to: string, comment: string): Promise<boolean> {
		const a = this.accountById(m.accountId);
		if (!a) return false;
		const addrs = to
			.split(/[,;]+/)
			.map((s) => s.trim())
			.filter(Boolean);
		if (!addrs.length) {
			new Notice("Power Desk: enter at least one address to forward to.");
			return false;
		}
		try {
			await forwardMessage(await this.graphTokenFor(a), m.id, comment.trim() || this.settings.mailSignature.trim() ? this.composeHtml(comment) : "", addrs);
			new Notice("Power Desk: forwarded.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** Where outgoing mail goes from: the preferred account when it can send,
	 *  else the first sending-capable account, else Power Assistant's
	 *  transport. Null still leaves the mail-app handoff. */
	mailSender(preferAccountId?: string): { label: string; send: (m: { to: string[]; cc?: string[]; subject: string; html: string }) => Promise<void> } | null {
		const preferred = preferAccountId ? this.accountById(preferAccountId) : null;
		const own =
			(preferred && preferred.refresh && preferred.grantedScope.includes("Mail.Send") ? preferred : null) ??
			this.settings.graphAccounts.find((a) => !!a.refresh && a.grantedScope.includes("Mail.Send")) ??
			null;
		if (own) {
			return { label: this.nameOf(own), send: async (m) => sendGraphMail(await this.graphTokenFor(own), m) };
		}
		const pa = this.assistantMailer();
		if (pa) return { label: "Power Assistant", send: pa };
		return null;
	}

	/** A message into the vault: dated note, readable text, source link. */
	async saveMailToNote(m: PCMail): Promise<void> {
		const body = await this.readMailBody(m);
		// mail has its own folder so it can be filed apart from event notes (and
		// into a protected folder); empty falls back to the calendar folder
		const folder = normalizePath(this.settings.mailNotesFolder.trim() || this.settings.notesFolder.trim() || "Calendar");
		await this.ensureFolder(folder);
		const path = normalizePath(`${folder}/${sanitizeName(`${keyOfMs(m.receivedMs)} ${m.subject}`)}.md`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.app.workspace.getLeaf(this.settings.notesInNewTab).openFile(existing);
			return;
		}
		const lines = [
			"---",
			`subject: ${JSON.stringify(m.subject)}`,
			`from: ${JSON.stringify(m.from)}`,
			`date: ${keyOfMs(m.receivedMs)}`,
			...(m.webLink ? [`source: ${JSON.stringify(m.webLink)}`] : []),
			"---",
			"",
			`# ${m.subject}`,
			"",
			body?.text ?? m.preview,
			"",
		];
		const f = await this.app.vault.create(path, lines.join("\n"));
		await this.app.workspace.getLeaf(this.settings.notesInNewTab).openFile(f);
		new Notice("Power Desk: mail saved to a note.");
	}

	async openMailView(): Promise<void> {
		const existing = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0];
		if (existing) {
			await this.app.workspace.revealLeaf(existing);
			return;
		}
		const leaf = this.app.workspace.getLeaf(true);
		await leaf.setViewState({ type: VIEW_TYPE_MAIL, active: true });
	}

	/** Power Assistant's own-mailbox mail transport when it is present and
	 *  connected; null otherwise. The mail-app handoff always works, so a
	 *  missing sibling costs a nicety, never the feature. */
	assistantMailer(): ((m: { to: string[]; subject: string; html: string }) => Promise<void>) | null {
		const plugs = (
			this.app as unknown as {
				plugins?: { plugins?: Record<string, { sendPageMail?: (m: { to: string[]; subject: string; html: string }) => Promise<void>; graphConnected?: () => boolean }> };
			}
		).plugins?.plugins;
		const pa = plugs?.["powerassistant"] ?? plugs?.["powercapture"];
		if (pa && typeof pa.sendPageMail === "function" && typeof pa.graphConnected === "function" && pa.graphConnected()) {
			return (m) => pa.sendPageMail!(m);
		}
		return null;
	}

	/** Recent mail as stripped text, for a sibling's rolling search window.
	 *
	 *  Bodies are fetched in batches and reduced to text before they leave, so
	 *  the caller holds a searchable corpus without the HTML weight and without
	 *  a note per message. `sinceMs` bounds the pull; `cap` bounds the batch so
	 *  one refresh cannot fetch a year of mail at once. Only messages already in
	 *  the mail list are considered, so this rides whatever the sync last saw. */
	async mailForIndex(sinceMs: number, cap = Math.min(5000, Math.max(50, this.settings.mailMaxMessages || 50))): Promise<{ id: string; from: string; subject: string; date: string; webLink?: string; text: string }[]> {
		const recent = this.allMail()
			.filter((m) => m.receivedMs >= sinceMs)
			.slice(0, cap);
		if (!recent.length) return [];
		await this.readMailBodies(recent, true);
		const out: { id: string; from: string; subject: string; date: string; webLink?: string; text: string }[] = [];
		for (const m of recent) {
			const body = this.bodyCache.get(m.id);
			if (!body) continue;
			out.push({
				id: m.id,
				from: m.fromAddress ? `${m.from} <${m.fromAddress}>` : m.from,
				subject: m.subject,
				date: new Date(m.receivedMs).toISOString(),
				webLink: m.webLink,
				text: body.text,
			});
		}
		return out;
	}

	/** Every mail folder across connected accounts, for a sibling's import
	 *  picker. Reads the cached tree, refreshing it if it has never loaded. */
	async mailFoldersForImport(): Promise<{ accountId: string; accountName: string; id: string; name: string; path: string }[]> {
		const out: { accountId: string; accountName: string; id: string; name: string; path: string }[] = [];
		for (const a of this.mailAccounts()) {
			if (!this.mailFolderCache.get(a.id)?.folders.length) await this.fetchFoldersFor(a);
			for (const row of this.folderTreeFor(a)) {
				out.push({
					accountId: a.id,
					accountName: this.nameOf(a),
					id: row.folder.id,
					name: row.folder.name,
					path: `${this.nameOf(a)} / ${row.folder.name}`,
				});
			}
		}
		return out;
	}

	/** One folder's messages with bodies, for a sibling to turn into notes.
	 *
	 *  Fetches the list deep (not the triage window), then batches the bodies
	 *  twenty at a time. Carries conversationId so the caller can collapse a
	 *  back-and-forth into a single exchange, and Outlook's own focused/other
	 *  verdict so it can filter on relevance without an AI call. */
	async mailFromFolder(
		accountId: string,
		folderId: string,
		opts: { sinceMs?: number; cap?: number } = {}
	): Promise<{ id: string; from: string; subject: string; date: string; webLink?: string; text: string; conversationId: string; focused: boolean; to: string }[]> {
		const a = this.accountById(accountId);
		if (!a) return [];
		const sinceMs = opts.sinceMs ?? Date.now() - 3650 * 86400000;
		const cap = Math.min(5000, Math.max(1, opts.cap ?? 2000));
		const token = await this.graphTokenFor(a);
		const raw = await this.mailboxGate(a.id, () => fetchFolderMessagesDeep(token, folderId, sinceMs, cap));
		const heads = raw
			.map((m) => {
				const o = m as Record<string, unknown>;
				const pc = graphMailToPC(o as GraphMailLike, a.id, this.nameOf(a), a.label);
				if (!pc) return null;
				return {
					pc,
					conversationId: String(o.conversationId ?? ""),
					focused: String(o.inferenceClassification ?? "focused").toLowerCase() !== "other",
					to: ((o.toRecipients as { emailAddress?: { name?: string; address?: string } }[] | undefined) ?? [])
						.map((r) => r.emailAddress?.name || r.emailAddress?.address || "")
						.filter(Boolean)
						.join(", "),
				};
			})
			.filter((x): x is NonNullable<typeof x> => x != null);
		if (!heads.length) return [];
		await this.readMailBodies(heads.map((h) => h.pc), true);
		const out: { id: string; from: string; subject: string; date: string; webLink?: string; text: string; conversationId: string; focused: boolean; to: string }[] = [];
		for (const h of heads) {
			const body = this.bodyCache.get(h.pc.id);
			out.push({
				id: h.pc.id,
				from: h.pc.fromAddress ? `${h.pc.from} <${h.pc.fromAddress}>` : h.pc.from,
				subject: h.pc.subject,
				date: new Date(h.pc.receivedMs).toISOString(),
				webLink: h.pc.webLink,
				text: body?.text ?? h.pc.preview,
				conversationId: h.conversationId,
				focused: h.focused,
				to: h.to,
			});
		}
		return out;
	}

	/* ---------------- transaction scanning ---------------- */


	/** True while a scan is running, so a timer tick that lands on a slow scan
	 *  queues nothing and simply waits for the next one. */
	private txnScanning = false;

	/** Power Assistant's transaction entry points, when it is installed and
	 *  current. Same config-borrowing rules as the meeting handoff: probe both
	 *  ids, feature-detect every method, never notice about a missing sibling. */
	assistantTxn(): {
		ready: () => boolean;
		select: (m: TxnMailLike[]) => TxnMailLike[];
		capture: (m: TxnMailLike & { html?: string; text?: string; webLink?: string; attachments?: string[] }) => Promise<number>;
	} | null {
		const plugs = (
			this.app as unknown as {
				plugins?: {
					plugins?: Record<
						string,
						{
							api?: {
								transactionsReady?: () => boolean;
								transactionSelect?: (m: TxnMailLike[]) => TxnMailLike[];
								captureTransaction?: (m: TxnMailLike & { html?: string; text?: string; webLink?: string; attachments?: string[] }) => Promise<number>;
							};
						}
					>;
				};
			}
		).plugins?.plugins;
		const api = plugs?.["powerassistant"]?.api ?? plugs?.["powercapture"]?.api;
		if (typeof api?.transactionsReady !== "function" || typeof api.transactionSelect !== "function" || typeof api.captureTransaction !== "function") return null;
		return { ready: api.transactionsReady.bind(api), select: api.transactionSelect.bind(api), capture: api.captureTransaction.bind(api) };
	}

	/** One pass over recent mail: refresh the lists, ask Power Assistant which
	 *  messages its rules want, then fetch a body and capture for each.
	 *
	 *  Bodies are only fetched for messages that already matched a rule, which
	 *  is why the selection happens on headers alone. A capture that throws is
	 *  logged and skipped rather than aborting the pass, so one malformed
	 *  message cannot stall every order behind it. */
	async scanForTransactions(force = false): Promise<number> {
		if (!this.settings.txnScan && !force) return 0;
		if (this.txnScanning) return 0;
		const txn = this.assistantTxn();
		if (!txn || !txn.ready()) return 0;
		this.txnScanning = true;
		try {
			this.ensureMail(false);
			const headers: TxnMailLike[] = this.allMail().map((m) => ({
				id: m.id,
				from: m.fromAddress ? `${m.from} <${m.fromAddress}>` : m.from,
				subject: m.subject,
				date: new Date(m.receivedMs).toISOString(),
				hasAttachments: !!m.hasAttachments,
			}));
			if (!headers.length) return 0;
			const wanted = txn.select(headers);
			if (!wanted.length) return 0;
			const byId = new Map(this.allMail().map((m) => [m.id, m]));
			// warm every matched body in ~one batch per twenty before reading them
			// back from the cache, instead of a round trip apiece
			const msgs = wanted.map((w) => byId.get(w.id)).filter((m): m is PCMail => !!m);
			await this.readMailBodies(msgs, true);
			let captured = 0;
			for (const w of wanted) {
				const msg = byId.get(w.id);
				if (!msg) continue;
				try {
					const body = await this.readMailBody(msg, true);
					if (!body) continue;
					// an invoice is often the attachment, not the message; save the
					// PDFs into the vault and let the sibling read their text
					const attachments: string[] = [];
					if (msg.hasAttachments) {
						for (const att of await this.mailAttachments(msg)) {
							if (!/\.pdf$/i.test(att.name)) continue;
							const path = await this.saveMailAttachment(msg, att);
							if (path) attachments.push(path);
						}
					}
					captured += await txn.capture({ ...w, html: body.html, text: body.text, webLink: msg.webLink, attachments });
				} catch (e) {
					console.warn("Power Desk: transaction capture failed for one message.", e);
				}
			}
			return captured;
		} finally {
			this.txnScanning = false;
		}
	}

	/* ---------------- Power Assistant handoff ---------------- */

	/** Power Assistant's meeting-note entry point when it is installed and
	 *  current; null otherwise. Config-borrowing rules apply: probe both ids,
	 *  feature-detect the method, never notice about a missing sibling. */
	assistantNewMeeting(): ((invite: Record<string, unknown>) => void) | null {
		const plugs = (this.app as unknown as { plugins?: { plugins?: Record<string, { api?: { newMeeting?: (invite: Record<string, unknown>) => void } }> } }).plugins
			?.plugins;
		const api = plugs?.["powerassistant"]?.api ?? plugs?.["powercapture"]?.api;
		return typeof api?.newMeeting === "function" ? api.newMeeting.bind(api) : null;
	}

	/** Hand an event to Power Assistant's New meeting dialog, prefilled. */
	captureMeeting(ev: PCEvent) {
		const open = this.assistantNewMeeting();
		if (!open) return;
		const span = eventDaySpan(ev);
		open({
			title: ev.title,
			date: span.startKey,
			when: ev.allDay ? "" : fmtEventRange(ev, this.settings.use24h),
			attendees: [...(ev.organizer ? [ev.organizer] : []), ...dedupePeople(ev.organizer, ev.attendees)],
			location: ev.location && !/^https?:\/\//i.test(ev.location) ? ev.location : "",
			teamsUrl: ev.joinUrl ?? "",
		});
	}

	/* ---------------- event notes ---------------- */

	eventNotePath(ev: PCEvent): string {
		const folder = normalizePath(this.settings.notesFolder.trim() || "Calendar");
		return normalizePath(`${folder}/${renderNoteName(this.settings.noteNameTemplate, ev, this.settings.use24h)}.md`);
	}

	noteExistsFor(ev: PCEvent): boolean {
		if (ev.notePath) return true; // the event IS a note
		return this.app.vault.getAbstractFileByPath(this.eventNotePath(ev)) instanceof TFile;
	}

	async openEventNote(ev: PCEvent) {
		if (ev.notePath) {
			const f = this.app.vault.getAbstractFileByPath(ev.notePath);
			if (f instanceof TFile) await this.app.workspace.getLeaf(this.settings.notesInNewTab).openFile(f);
			return;
		}
		const path = this.eventNotePath(ev);
		let file = this.app.vault.getAbstractFileByPath(path);
		if (!(file instanceof TFile)) {
			const folder = normalizePath(this.settings.notesFolder.trim() || "Calendar");
			await this.ensureFolder(folder);
			file = await this.app.vault.create(path, eventNoteContent(ev, this.settings.use24h));
			this.notify(); // note dots update
		}
		if (file instanceof TFile) await this.app.workspace.getLeaf(this.settings.notesInNewTab).openFile(file);
	}

	private async ensureFolder(path: string) {
		const p = normalizePath(path);
		if (this.app.vault.getAbstractFileByPath(p) instanceof TFolder) return;
		await this.app.vault.createFolder(p).catch(() => {});
	}

	/** Where person pages live: own setting first, Power Assistant's People
	 *  folder as the fallback, a plain "People" folder when neither exists. */
	personFolderPath(): string {
		const own = this.settings.peopleFolder.trim();
		if (own) return normalizePath(own);
		const plugs = (this.app as unknown as { plugins?: { plugins?: Record<string, { settings?: { peopleFolder?: string; outputFolder?: string } }> } }).plugins?.plugins;
		const pa = (plugs?.["powerassistant"] ?? plugs?.["powercapture"])?.settings;
		if (pa?.peopleFolder?.trim()) return normalizePath(pa.peopleFolder.trim());
		if (pa?.outputFolder?.trim()) return normalizePath(`${pa.outputFolder.trim()}/People`);
		return "People";
	}

	/** Open (or start) the person page behind an attendee name, so the people
	 *  on your meetings connect to their hubs in one click. */
	async openPersonPage(name: string) {
		const folder = this.personFolderPath();
		const path = normalizePath(`${folder}/${sanitizeName(name)}.md`);
		let f = this.app.vault.getAbstractFileByPath(path);
		if (!(f instanceof TFile)) {
			await this.ensureFolder(folder);
			f = await this.app.vault.create(path, `# ${sanitizeName(name)}\n`);
		}
		if (f instanceof TFile) await this.app.workspace.getLeaf(this.settings.notesInNewTab).openFile(f);
	}

	/** Drop a ready-made Bases table over the event-notes folder (Power Bases'
	 *  table when it is installed, the core one otherwise) and open it. */
	async createEventsBase() {
		const folder = normalizePath(this.settings.notesFolder.trim() || "Calendar");
		const path = normalizePath(`${folder}/Events.base`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.app.workspace.getLeaf(false).openFile(existing);
			return;
		}
		const plugs = (this.app as unknown as { plugins?: { plugins?: Record<string, unknown> } }).plugins?.plugins;
		const viewType = plugs?.["powerbases"] ? "powerbases-table" : "table";
		const yaml = [
			"filters:",
			"  and:",
			`    - file.inFolder("${folder}")`,
			'    - file.ext == "md"',
			"views:",
			`  - type: ${viewType}`,
			"    name: Events",
			"    order:",
			"      - file.name",
			"      - note.date",
			"      - note.time",
			"      - note.calendar",
			"      - note.location",
			"      - note.organizer",
			"",
		].join("\n");
		await this.ensureFolder(folder);
		const f = await this.app.vault.create(path, yaml);
		new Notice("Power Desk: events base created.");
		await this.app.workspace.getLeaf(false).openFile(f);
	}

	/* ---------------- view plumbing ---------------- */

	async openCalendarView(): Promise<PowerCalendarView | null> {
		const existing = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0];
		if (existing) {
			await this.app.workspace.revealLeaf(existing);
			return existing.view instanceof PowerCalendarView ? existing.view : null;
		}
		const leaf = this.app.workspace.getLeaf(true);
		await leaf.setViewState({ type: VIEW_TYPE, active: true });
		return leaf.view instanceof PowerCalendarView ? leaf.view : null;
	}

	openOwnSettings() {
		const setting = (this.app as unknown as { setting?: { open: () => void; openTabById: (id: string) => void } }).setting;
		setting?.open();
		setting?.openTabById(this.manifest.id);
	}
}

/** The note a calendar event creates: enough frontmatter to query from Bases,
 *  attendees as links so people pages connect, and the body left to the user. */
function eventNoteContent(ev: PCEvent, use24h: boolean): string {
	const s = eventDaySpan(ev);
	const lines = ["---", `event: ${JSON.stringify(ev.title)}`, `date: ${s.startKey}`];
	if (!ev.allDay) lines.push(`time: ${JSON.stringify(fmtEventRange(ev, use24h))}`);
	if (ev.calendarName) lines.push(`calendar: ${JSON.stringify(ev.calendarName)}`);
	if (ev.location) lines.push(`location: ${JSON.stringify(ev.location)}`);
	if (ev.organizer) lines.push(`organizer: ${JSON.stringify(ev.organizer)}`);
	if (ev.attendees?.length) {
		lines.push("attendees:");
		for (const a of ev.attendees) lines.push(`  - "[[${a.replace(/[[\]|#^\\]/g, " ").replace(/\s+/g, " ").trim()}]]"`);
	}
	if (ev.joinUrl) lines.push(`join: ${JSON.stringify(ev.joinUrl)}`);
	if (ev.url) lines.push(`source: ${JSON.stringify(ev.url)}`);
	lines.push("---", "", `# ${ev.title}`, "");
	if (ev.description) lines.push(ev.description.trim(), "");
	return lines.join("\n");
}

const sleep = (ms: number) => new Promise<void>((r) => window.setTimeout(r, ms));

/** The agenda-in-a-note block. Lives as a render child so it follows the
 *  plugin's data (a fetch landing repaints it) and unhooks when the block
 *  leaves the DOM. Renders as a block widget, so Live Preview keeps it while
 *  the note stays editable. */
class AgendaBlock extends MarkdownRenderChild {
	private readonly onData = () => this.render();

	constructor(
		containerEl: HTMLElement,
		private plugin: PowerDeskPlugin,
		private cfg: { date: string | null; days: number },
		private sourcePath: string
	) {
		super(containerEl);
	}

	onload() {
		this.plugin.listeners.add(this.onData);
		this.render();
	}

	onunload() {
		this.plugin.listeners.delete(this.onData);
	}

	private startKey(): string {
		if (this.cfg.date) return this.cfg.date;
		const base = this.sourcePath.split("/").pop() ?? "";
		return dateKeyOf(base) ?? keyOfDate(new Date());
	}

	private render() {
		const el = this.containerEl;
		el.empty();
		el.addClass("pcal-embed");
		const s = this.plugin.settings;
		const from = this.startKey();
		const to = addDays(from, this.cfg.days - 1);
		this.plugin.ensureWindow(from, to, false);
		const openAt = (key: string) => void this.plugin.openCalendarView().then((v) => v?.goDay(key));
		const title = el.createDiv({ cls: "pcal-embed-title", text: this.cfg.days === 1 ? fmtDayHeading(from) : `${fmtDayShort(from)} - ${fmtDayShort(to, true)}` });
		title.addEventListener("click", () => openAt(from));
		const groups = groupByDay(this.plugin.eventsForWindow(from, to), from, to);
		if (!groups.length) {
			el.createDiv({ cls: "pcal-embed-empty", text: this.plugin.sources().length ? "Nothing scheduled." : "No calendar sources connected." });
			return;
		}
		for (const g of groups) {
			if (this.cfg.days > 1) {
				const head = el.createDiv({ cls: "pcal-agenda-head", text: fmtDayHeading(g.key) });
				head.addEventListener("click", () => openAt(g.key));
			}
			for (const ev of g.events) {
				const row = el.createDiv("pcal-agenda-row");
				row.style.setProperty("--pcal-ev-color", ev.color ?? "var(--interactive-accent)");
				row.toggleClass("is-declined", !!ev.declined);
				row.createDiv("pcal-agenda-dot");
				row.createDiv({ cls: "pcal-agenda-time", text: ev.allDay ? "All day" : fmtTimeOfMs(ev.startMs, s.use24h) });
				const main = row.createDiv("pcal-agenda-main");
				main.createDiv({ cls: "pcal-agenda-title", text: ev.title });
				if (ev.location) main.createDiv({ cls: "pcal-agenda-sub", text: ev.location });
				if (ev.joinUrl) {
					const join = row.createEl("button", { cls: "pcal-icon-btn pcal-join-btn", attr: { "aria-label": "Join meeting" } });
					setIcon(join, "video");
					join.addEventListener("click", (e) => {
						e.stopPropagation();
						window.open(ev.joinUrl, "_blank");
					});
				}
				row.addEventListener("click", () => openAt(g.key));
			}
		}
	}
}

/** Jump to any upcoming (or last week's) event by fuzzy title, day, calendar,
 *  or attendee. Enter lands the calendar on its day. */
class EventFindModal extends FuzzySuggestModal<PCEvent> {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
		this.setPlaceholder("Find an event...");
	}

	getItems(): PCEvent[] {
		const today = keyOfDate(new Date());
		const from = addDays(today, -7);
		const to = addDays(today, 60);
		this.plugin.ensureWindow(from, to, false);
		return this.plugin.eventsForWindow(from, to).sort((a, b) => a.startMs - b.startMs);
	}

	getItemText(ev: PCEvent): string {
		const span = eventDaySpan(ev);
		const when = ev.allDay ? "" : ` ${fmtTimeOfMs(ev.startMs, this.plugin.settings.use24h, true)}`;
		return `${ev.title}  ${fmtDayShort(span.startKey, true)}${when}  ${ev.calendarName ?? ""}  ${(ev.attendees ?? []).join(" ")}`;
	}

	onChooseItem(ev: PCEvent): void {
		void this.plugin.openCalendarView().then((v) => v?.goDay(eventDaySpan(ev).startKey));
	}
}

/** The triage inbox: recent mail across every mail-enabled account, a
 *  text-first reading pane, and the two bridges that make mail belong in a
 *  calendar plugin: Make event and Save to note. Deliberately not a mail
 *  client: no folders, no rules, no mailbox-wide search. */
class MailView extends ItemView {
	private foldersEl!: HTMLElement;
	private listEl!: HTMLElement;
	private readEl!: HTMLElement;
	private titleTextEl!: HTMLElement;
	private refreshBtn!: HTMLElement;
	private backBtn!: HTMLElement;
	private foldToggleBtn!: HTMLElement;
	private searchRowEl!: HTMLElement;
	private searchInputEl!: HTMLInputElement;
	/** Null = the unified All-inboxes list. */
	private folderSel: { accountId: string; folderId: string; name: string } | null = null;
	/** Narrow layouts drill one screen at a time, like the phone Mail app. */
	private drill = false;
	private screen: "folders" | "list" | "read" = "list";
	private selected: PCMail | null = null;
	private selectedBody: { text: string; html?: string; toLine: string } | null = null;
	private selectedAtts: MailAttachment[] | null = null;
	/** Inline images resolved from cid: references to data urls. */
	private inlineCids: Map<string, string> | null = null;
	/** The favorite row a drag started from, for drop reordering. */
	private favDragIdx: number | null = null;
	/** Ctrl and Shift clicks build a bulk selection for the toolbar actions. */
	private multiSel = new Set<string>();
	private selAnchorId: string | null = null;
	/** The list as last rendered, in order, for shift ranges and bulk lookup. */
	private lastList: PCMail[] = [];
	/** Set when a folder is entered: the next render points the reading pane
	 *  at the list's first message. */
	private autoSelectPending = false;
	/** The reply box's state lives on the view, so a refresh or the body
	 *  arriving cannot wipe a half-typed draft. */
	private mailToolBtns: HTMLElement[] = [];
	private readTimer: number | null = null;
	/** List-header state: one extra filter beside Unread, and the sort. */
	private extraFilter: "none" | "priority" | "flagged" | "tome" | "attachments" = "none";
	/** Select mode shows every checkbox unchecked, ready for picking. */
	private selectMode = false;
	private sortBy: "date" | "from" | "subject" = "date";
	private sortAsc = false;
	private renderQueued = false;
	private readonly onData = () => this.queueRender();

	constructor(
		leaf: WorkspaceLeaf,
		private plugin: PowerDeskPlugin
	) {
		super(leaf);
		this.scope = new Scope(this.app.scope);
		const guard = (evt: KeyboardEvent, cb: () => void) => {
			const t = evt.target as HTMLElement | null;
			if (t && (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t.isContentEditable)) return true;
			cb();
			return false;
		};
		(this.scope as Scope).register([], "Delete", (evt) => guard(evt, () => this.deleteSelection(false)));
		(this.scope as Scope).register(["Shift"], "Delete", (evt) => guard(evt, () => this.deleteSelection(true)));
	}

	/** Del deletes the selection into Deleted Items; Shift+Del deletes it
	 *  permanently, behind a confirm because nothing brings that back. */
	private deleteSelection(permanent: boolean) {
		this.deleteMessages(this.multiTargets(), permanent);
	}

	private deleteMessages(targets: PCMail[], permanent: boolean) {
		if (!targets.length) return;
		const finish = () => {
			if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
			this.multiSel.clear();
			this.render();
		};
		if (!permanent) {
			for (const t of targets) void this.plugin.deleteMail(t);
			finish();
			return;
		}
		new ConfirmModal(
			this.app,
			targets.length > 1 ? `Permanently delete ${targets.length} messages?` : `Permanently delete "${targets[0].subject}"?`,
			"They skip Deleted Items and cannot be recovered from the mailbox.",
			"Delete permanently",
			() => {
				for (const t of targets) void this.plugin.permanentDeleteMail(t);
				finish();
			}
		).open();
	}

	getViewType(): string {
		return VIEW_TYPE_MAIL;
	}

	getDisplayText(): string {
		return "Inbox";
	}

	getIcon(): string {
		return "mail";
	}

	async onOpen() {
		this.plugin.listeners.add(this.onData);
		const root = this.contentEl;
		root.empty();
		root.addClass("pcal-mail-root");
		const header = root.createDiv("pcal-mail-header");
		this.backBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Back" } });
		setIcon(this.backBtn, "chevron-left");
		this.backBtn.addEventListener("click", () => {
			this.screen = this.screen === "read" ? "list" : "folders";
			this.render();
		});
		this.foldToggleBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Folders" } });
		setIcon(this.foldToggleBtn, "panel-left");
		this.foldToggleBtn.addEventListener("click", () => root.toggleClass("pcal-folders-hidden", !root.hasClass("pcal-folders-hidden")));
		const tools = header.createDiv("pcal-mail-tools");
		const newMail = tools.createEl("button", { cls: "pcal-new-btn", text: "New mail" });
		newMail.addEventListener("click", () => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open());
		// the Outlook ribbon pattern: labeled actions on the selection, which
		// is the bulk set when Ctrl or Shift built one
		const tool = (label: string, icon: string, cb: () => void) => {
			const b = tools.createEl("button", { cls: "pcal-icon-btn pcal-mail-tool", attr: { "aria-label": label } });
			const ic = b.createSpan();
			setIcon(ic, icon);
			b.createSpan({ cls: "pcal-mail-tool-label", text: label });
			b.addEventListener("click", () => {
				if (!this.selected && !this.multiSel.size) {
					new Notice("Power Desk: select a message first.");
					return;
				}
				cb();
			});
			this.mailToolBtns.push(b);
		};
		const single = (): PCMail | null => {
			if (this.multiSel.size > 1) {
				new Notice("Power Desk: this works on a single message; click one without Ctrl.");
				return null;
			}
			return this.multiTargets()[0] ?? null;
		};
		tool("Delete", "trash-2", () => {
			const targets = this.multiTargets();
			for (const t of targets) void this.plugin.deleteMail(t);
			if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
			this.multiSel.clear();
			this.render();
		});
		tool("Reply", "reply", () => {
			const m = single();
			if (m) void this.selectAndReply(m);
		});
		tool("Reply all", "reply-all", () => {
			const m = single();
			if (m) void this.selectAndReply(m, true);
		});
		tool("Forward", "forward", () => {
			const m = single();
			if (m) new RichComposeModal(this.app, this.plugin, { mode: "forward", mail: m }).open();
		});
		tool("Read / Unread", "mail-open", () => {
			const targets = this.multiTargets();
			const anyUnread = targets.some((t) => t.unread);
			for (const t of targets) void this.plugin.setMailRead(t, anyUnread);
		});
		const right = header.createDiv("pcal-header-right");
		const searchBtn = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Search mail" } });
		setIcon(searchBtn, "search");
		searchBtn.addEventListener("click", () => {
			const open = this.searchRowEl.isShown();
			if (open) {
				this.searchRowEl.hide();
				this.plugin.clearMailSearch();
			} else {
				this.searchRowEl.show();
				this.searchInputEl.focus();
			}
		});
		this.refreshBtn = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Refresh" } });
		setIcon(this.refreshBtn, "refresh-cw");
		this.refreshBtn.addEventListener("click", () => this.plugin.refreshMailAll(true));
		const splitR = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Split right" } });
		setIcon(splitR, "separator-vertical");
		splitR.addEventListener("click", () => void this.app.workspace.duplicateLeaf(this.leaf, "split", "vertical"));
		const splitD = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Split down" } });
		setIcon(splitD, "separator-horizontal");
		splitD.addEventListener("click", () => void this.app.workspace.duplicateLeaf(this.leaf, "split", "horizontal"));
		this.searchRowEl = root.createDiv("pcal-mail-searchrow");
		this.searchRowEl.hide();
		this.searchInputEl = this.searchRowEl.createEl("input", { attr: { type: "search", placeholder: "Search the whole mailbox..." } });
		this.searchInputEl.addEventListener("keydown", (e) => {
			if (e.key === "Enter") void this.plugin.runMailSearch(this.searchInputEl.value);
			if (e.key === "Escape") {
				e.stopPropagation();
				this.searchRowEl.hide();
				this.plugin.clearMailSearch();
			}
		});
		const searchGo = this.searchRowEl.createEl("button", { text: "Search", cls: "mod-cta" });
		searchGo.addEventListener("click", () => void this.plugin.runMailSearch(this.searchInputEl.value));
		const searchClear = this.searchRowEl.createEl("button", { text: "Clear" });
		searchClear.addEventListener("click", () => {
			this.searchInputEl.value = "";
			this.searchRowEl.hide();
			this.plugin.clearMailSearch();
		});
		const body = root.createDiv("pcal-mail-body");
		this.foldersEl = body.createDiv("pcal-mail-folders");
		const split1 = Platform.isPhone ? null : body.createDiv("pcal-mail-splitter");
		// the list column: an Outlook-style header (folder name, select, jump,
		// filter, sort) over the rows
		const listWrap = body.createDiv("pcal-mail-listwrap");
		const lh = listWrap.createDiv("pcal-mail-listhead");
		this.titleTextEl = lh.createSpan({ cls: "pcal-mail-listtitle", text: "All inboxes" });
		const lhBtns = lh.createDiv("pcal-mail-listhead-btns");
		const hbtn = (icon: string, label: string, cb: (e: MouseEvent) => void): HTMLElement => {
			const b = lhBtns.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
			setIcon(b, icon);
			b.addEventListener("click", cb);
			return b;
		};
		const selBtn = hbtn("copy-check", "Select messages, then tick the ones you want", () => {
			this.selectMode = !this.selectMode;
			if (!this.selectMode) this.multiSel.clear();
			selBtn.toggleClass("is-active", this.selectMode);
			this.render();
		});
		hbtn("arrow-down", "Jump to a date", (e) => {
			const menu = new Menu();
			const jump = (label: string, daysBack: number) =>
				menu.addItem((i) =>
					i.setTitle(label).onClick(() => {
						const targetMs = msOfKey(addDays(keyOfDate(new Date()), -daysBack)) + 86399999;
						const idx = this.lastList.findIndex((m) => m.receivedMs <= targetMs);
						const rows = this.listEl.querySelectorAll(".pcal-mail-row");
						if (idx >= 0 && rows[idx]) (rows[idx] as HTMLElement).scrollIntoView({ block: "start" });
					})
				);
			jump("Today", 0);
			jump("Yesterday", 1);
			jump("Last week", 7);
			jump("Last month", 30);
			menu.showAtMouseEvent(e);
		});
		hbtn("filter", "Filter the list", (e) => {
			const s = this.plugin.settings;
			const menu = new Menu();
			menu.addItem((i) =>
				i.setTitle("All")
					.setChecked(!s.mailUnreadOnly && this.extraFilter === "none")
					.onClick(() => {
						s.mailUnreadOnly = false;
						this.extraFilter = "none";
						this.plugin.queueSave();
						this.render();
					})
			);
			menu.addItem((i) =>
				i.setTitle("Unread")
					.setChecked(s.mailUnreadOnly)
					.onClick(() => {
						s.mailUnreadOnly = !s.mailUnreadOnly;
						this.plugin.queueSave();
						this.render();
					})
			);
			const extra = (title: string, v: "priority" | "flagged" | "tome" | "attachments") =>
				menu.addItem((i) =>
					i.setTitle(title)
						.setChecked(this.extraFilter === v)
						.onClick(() => {
							this.extraFilter = this.extraFilter === v ? "none" : v;
							this.render();
						})
				);
			extra("High priority", "priority");
			extra("Flagged", "flagged");
			extra("To me", "tome");
			extra("Has attachments", "attachments");
			menu.showAtMouseEvent(e);
		});
		hbtn("arrow-up-down", "Sort the list", (e) => {
			const menu = new Menu();
			const by = (label: string, v: "date" | "from" | "subject") =>
				menu.addItem((i) =>
					i.setTitle(label)
						.setChecked(this.sortBy === v)
						.onClick(() => {
							this.sortBy = v;
							this.render();
						})
				);
			by("Date", "date");
			by("From", "from");
			by("Subject", "subject");
			menu.addSeparator();
			menu.addItem((i) =>
				i.setTitle("Newest on top")
					.setChecked(!this.sortAsc)
					.onClick(() => {
						this.sortAsc = false;
						this.render();
					})
			);
			menu.addItem((i) =>
				i.setTitle("Oldest on top")
					.setChecked(this.sortAsc)
					.onClick(() => {
						this.sortAsc = true;
						this.render();
					})
			);
			menu.showAtMouseEvent(e);
		});
		this.listEl = listWrap.createDiv("pcal-mail-list");
		const split2 = Platform.isPhone ? null : body.createDiv("pcal-mail-splitter");
		this.readEl = body.createDiv("pcal-mail-read");
		// pane widths drag on the seams and stick per device, not per vault
		const store = this.app as unknown as { loadLocalStorage: (k: string) => unknown; saveLocalStorage: (k: string, v: unknown) => void };
		const wireSplitter = (el: HTMLElement | null, pane: HTMLElement, storeKey: string, min: number, max: number) => {
			if (!el) return;
			const saved = Number(store.loadLocalStorage(storeKey) ?? 0);
			if (saved >= min && saved <= max) {
				pane.style.flex = "0 0 auto";
				pane.style.width = `${saved}px`;
			}
			el.addEventListener("pointerdown", (e) => {
				e.preventDefault();
				const startX = e.clientX;
				const startW = pane.getBoundingClientRect().width;
				pane.style.flex = "0 0 auto";
				const move = (ev: PointerEvent) => {
					pane.style.width = `${Math.min(max, Math.max(min, startW + (ev.clientX - startX)))}px`;
				};
				const up = () => {
					document.removeEventListener("pointermove", move);
					document.removeEventListener("pointerup", up);
					store.saveLocalStorage(storeKey, String(Math.round(pane.getBoundingClientRect().width)));
				};
				document.addEventListener("pointermove", move);
				document.addEventListener("pointerup", up);
			});
		};
		wireSplitter(split1, this.foldersEl, "pdesk-mail-folders-w", 140, 480);
		wireSplitter(split2, listWrap, "pdesk-mail-list-w", 200, 720);
		this.plugin.ensureMail(false);
		this.plugin.ensureMailFolders(false);
		// the drill follows the CONTAINER's width, so a narrow desktop pane
		// gets the phone treatment too and a rotated tablet leaves it
		const measure = () => {
			const w = this.contentEl.clientWidth;
			const drill = w > 0 && w < 640;
			if (drill !== this.drill) {
				this.drill = drill;
				if (drill) this.screen = this.selected ? "read" : "list";
				this.queueRender();
			}
		};
		const ro = new ResizeObserver(measure);
		ro.observe(this.contentEl);
		this.register(() => ro.disconnect());
		measure();
		this.render();
	}

	async onClose() {
		this.plugin.listeners.delete(this.onData);
		this.clearReadTimer();
	}

	private clearReadTimer() {
		if (this.readTimer != null) {
			window.clearTimeout(this.readTimer);
			this.readTimer = null;
		}
	}

	/** Whether selecting may mark read right now: the unread filter and the
	 *  Unread search folder can both pin everything unread until marked by
	 *  hand. */
	private autoReadSuppressed(): boolean {
		const s = this.plugin.settings;
		return (s.mailUnreadOnly || this.folderSel?.folderId === UNREAD_FOLDER) && s.unreadFilterKeepsUnread;
	}

	private queueRender() {
		if (this.renderQueued) return;
		this.renderQueued = true;
		requestAnimationFrame(() => {
			this.renderQueued = false;
			this.render();
		});
	}

	private render() {
		if (!this.listEl || !this.listEl.isConnected) return;
		const s = this.plugin.settings;
		const todayKey = keyOfDate(new Date());
		if (this.drill && this.screen === "read" && !this.selected) this.screen = "list";
		const root = this.contentEl;
		root.toggleClass("pcal-mail-drill", this.drill);
		root.toggleClass("pcal-screen-folders", this.drill && this.screen === "folders");
		root.toggleClass("pcal-screen-list", this.drill && this.screen === "list");
		root.toggleClass("pcal-screen-read", this.drill && this.screen === "read");
		this.backBtn.toggle(this.drill && this.screen !== "folders");
		this.foldToggleBtn.toggle(!this.drill);
		this.refreshBtn.toggleClass("is-loading", this.plugin.anyMailInFlight());
		this.listEl.empty();
		const accounts = this.plugin.mailAccounts();
		if (!accounts.length) {
			const empty = this.listEl.createDiv("pcal-empty");
			empty.createDiv({ text: "No mail-enabled account. Reconnect a Microsoft account in settings to grant mail access." });
			const b = empty.createEl("button", { text: "Open settings", cls: "mod-cta" });
			b.addEventListener("click", () => this.plugin.openOwnSettings());
			this.readEl.empty();
			return;
		}
		const colorOf = new Map(accounts.map((a, i) => [a.id, paletteColor(i)]));
		this.renderFolders(accounts, colorOf);
		const sel = this.folderSel;
		if (sel && !this.plugin.accountById(sel.accountId)) this.folderSel = null;
		const search = this.plugin.mailSearchState();
		this.titleTextEl.setText(
			search ? `Search: ${search.query}` : this.drill && this.screen === "folders" ? "Mailboxes" : this.folderSel ? this.folderSel.name : "All inboxes"
		);
		if (!search && this.folderSel) this.plugin.ensureFolderMail(this.folderSel.accountId, this.folderSel.folderId, false);
		const source = search ? search.results : this.folderSel ? this.plugin.folderMail(this.folderSel.accountId, this.folderSel.folderId) : this.plugin.allMail();
		// the unread filter keeps the currently open message visible even once
		// it reads, so it cannot vanish mid-read
		const mail = source.filter((m) => !s.mailUnreadOnly || m.unread || m.id === this.selected?.id);
		void this.plugin.prefetchBodies(mail.slice(0, 8));
		// entering a folder points the reading pane at its first message, so
		// the body beside the list always belongs to it; a selection already
		// in the list stays put, the phone drill has no pane to match, and a
		// search supersedes the folder entry
		if (this.autoSelectPending) {
			if (this.drill || search || mail.some((x) => x.id === this.selected?.id)) this.autoSelectPending = false;
			else if (mail.length) {
				this.autoSelectPending = false;
				void this.select(mail[0], true);
				return;
			} else if (!(this.folderSel ? this.plugin.folderMailInFlight(this.folderSel.accountId, this.folderSel.folderId) : this.plugin.anyMailInFlight())) {
				this.autoSelectPending = false;
				this.selected = null;
				this.selectedBody = null;
			}
		}
		if (search?.error) this.listEl.createDiv({ cls: "pcal-mail-error", text: search.error });
		else if (this.folderSel && !search) {
			const err = this.plugin.folderMailError(this.folderSel.accountId, this.folderSel.folderId);
			if (err) this.listEl.createDiv({ cls: "pcal-mail-error", text: err });
		} else if (!search) {
			for (const err of this.plugin.mailErrors()) this.listEl.createDiv({ cls: "pcal-mail-error", text: err });
		}
		if (!mail.length)
			this.listEl.createDiv({
				cls: "pcal-embed-empty",
				text: search ? (search.inFlight ? "Searching..." : "No matches.") : this.plugin.anyMailInFlight() ? "Loading..." : s.mailUnreadOnly ? "No unread mail." : this.folderSel ? "Nothing here." : "Inbox zero.",
			});
		// search results and the Unread folder both tag rows with their source folder
		const showTags = !!search || this.folderSel?.folderId === UNREAD_FOLDER;
		const nameMaps = showTags ? new Map(accounts.map((acc) => [acc.id, this.plugin.folderNamesFor(acc.id)])) : null;
		let shown = mail;
		if (this.extraFilter === "attachments") shown = shown.filter((m) => m.hasAttachments);
		else if (this.extraFilter === "priority") shown = shown.filter((m) => m.priority);
		else if (this.extraFilter === "flagged") shown = shown.filter((m) => m.flagged);
		else if (this.extraFilter === "tome") shown = shown.filter((m) => m.toMe);
		if (this.sortBy !== "date" || this.sortAsc)
			shown = [...shown].sort((x, y) => {
				const c = this.sortBy === "from" ? x.from.localeCompare(y.from) : this.sortBy === "subject" ? x.subject.localeCompare(y.subject) : x.receivedMs - y.receivedMs;
				return this.sortAsc ? c : -c;
			});
		this.lastList = shown;
		this.listEl.toggleClass("has-multisel", this.multiSel.size > 0);
		this.listEl.toggleClass("select-mode", this.selectMode);
		for (const m of shown) {
			const row = this.listEl.createDiv("pcal-mail-row");
			row.toggleClass("is-unread", m.unread);
			row.toggleClass("is-selected", this.selected?.id === m.id);
			row.toggleClass("is-multisel", this.multiSel.has(m.id));
			const av = row.createDiv("pcal-mail-avatar");
			av.style.backgroundColor = avatarColor(m.from || m.fromAddress);
			av.setText(avatarInitials(m.from || m.fromAddress));
			av.setAttribute("aria-label", m.accountLabel);
			// the Outlook checkbox: over the avatar on hover, or always while a
			// selection is active; clicking it never changes the reading pane
			const check = row.createDiv("pcal-mail-check");
			setIcon(check, this.multiSel.has(m.id) ? "check-square" : "square");
			check.addEventListener("click", (e) => {
				e.stopPropagation();
				if (this.multiSel.has(m.id)) this.multiSel.delete(m.id);
				else this.multiSel.add(m.id);
				this.selAnchorId = m.id;
				this.render();
			});
			const mid = row.createDiv("pcal-mail-mid");
			const top = mid.createDiv("pcal-mail-top");
			top.createSpan({ cls: "pcal-mail-from", text: m.from });
			const srcFolder = nameMaps && m.folderId ? nameMaps.get(m.accountId)?.get(m.folderId) : null;
			if (srcFolder) top.createSpan({ cls: "pcal-mail-tag", text: srcFolder });
			top.createSpan({ cls: "pcal-mail-time", text: fmtMailTime(m.receivedMs, todayKey, s.use24h) });
			const subj = mid.createDiv("pcal-mail-subject");
			if (m.hasAttachments) {
				const clip = subj.createSpan("pcal-mail-clip");
				setIcon(clip, "paperclip");
			}
			subj.createSpan({ text: m.subject });
			mid.createDiv({ cls: "pcal-mail-preview", text: m.preview });
			const actions = row.createDiv("pcal-mail-actions");
			const act = (icon: string, label: string, cb: () => void) => {
				const b = actions.createEl("button", { cls: "pcal-mail-act", attr: { "aria-label": label } });
				setIcon(b, icon);
				b.addEventListener("click", (e) => {
					e.stopPropagation();
					cb();
				});
			};
			act("reply-all", "Reply all", () => void this.selectAndReply(m, true));
			act(m.unread ? "mail-open" : "mail", m.unread ? "Mark read" : "Mark unread", () => void this.plugin.setMailRead(m, m.unread));
			act("trash-2", "Delete", () => {
				void this.plugin.deleteMail(m);
				if (this.selected?.id === m.id) {
					this.selected = null;
					this.renderReading();
				}
			});
			row.addEventListener("click", (e) => {
				// Ctrl toggles membership, Shift takes the range from the last
				// anchor; a plain click clears the bulk set and reads normally
				if (e.ctrlKey || e.metaKey) {
					if (this.multiSel.has(m.id)) this.multiSel.delete(m.id);
					else this.multiSel.add(m.id);
					this.selAnchorId = m.id;
					row.toggleClass("is-multisel", this.multiSel.has(m.id));
					return;
				}
				if (e.shiftKey) {
					const anchor = this.selAnchorId ?? this.selected?.id ?? null;
					const ai = anchor ? this.lastList.findIndex((x) => x.id === anchor) : -1;
					const bi = this.lastList.findIndex((x) => x.id === m.id);
					if (ai >= 0 && bi >= 0) {
						this.multiSel = new Set(this.lastList.slice(Math.min(ai, bi), Math.max(ai, bi) + 1).map((x) => x.id));
						this.render();
					}
					return;
				}
				if (this.selectMode) {
					// in select mode the whole row is a checkbox
					if (this.multiSel.has(m.id)) this.multiSel.delete(m.id);
					else this.multiSel.add(m.id);
					this.selAnchorId = m.id;
					this.render();
					return;
				}
				this.multiSel.clear();
				this.selAnchorId = m.id;
				void this.select(m);
			});
			row.addEventListener("contextmenu", (e) => {
				e.preventDefault();
				// the menu acts on the whole selection when this row is in it
				const targets = this.multiSel.has(m.id) && this.multiSel.size ? this.multiTargets() : [m];
				const plural = targets.length > 1 ? ` ${targets.length} messages` : "";
				const done = (clear: boolean) => {
					if (clear) {
						if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
						this.multiSel.clear();
					}
					this.render();
				};
				const menu = new Menu();
				menu.addItem((i) => i.setTitle(`Delete${plural}`).onClick(() => this.deleteMessages(targets, false)));
				menu.addItem((i) => i.setTitle(`Delete${plural} permanently`).onClick(() => this.deleteMessages(targets, true)));
				menu.addItem((i) =>
					i.setTitle(`Archive${plural}`).onClick(() => {
						for (const t of targets) void this.plugin.archiveMail(t);
						done(true);
					})
				);
				menu.addItem((i) =>
					i.setTitle(`Mark${plural} read`).onClick(() => {
						for (const t of targets) void this.plugin.setMailRead(t, true);
						done(false);
					})
				);
				menu.addItem((i) =>
					i.setTitle(`Mark${plural} unread`).onClick(() => {
						for (const t of targets) void this.plugin.setMailRead(t, false);
						done(false);
					})
				);
				if (this.multiSel.size) {
					menu.addSeparator();
					menu.addItem((i) =>
						i.setTitle("Clear selection").onClick(() => {
							this.multiSel.clear();
							this.render();
						})
					);
				}
				menu.showAtMouseEvent(e);
			});
		}
		this.renderReading();
	}

	/** What the toolbar acts on: the bulk selection when one exists, else the
	 *  message in the reading pane. */
	private multiTargets(): PCMail[] {
		if (this.multiSel.size) {
			const byId = new Map(this.lastList.map((x) => [x.id, x]));
			return [...this.multiSel].map((id) => byId.get(id)).filter((x): x is PCMail => !!x);
		}
		return this.selected ? [this.selected] : [];
	}

	/** The folder pane: each account as a group, its tree beneath, unread
	 *  counts riding along; All inboxes on top returns to the unified list. */
	private renderFolders(accounts: GraphAccount[], colorOf: Map<string, string>) {
		const host = this.foldersEl;
		if (!host) return;
		host.empty();
		const allRow = host.createDiv("pcal-folder-row");
		allRow.toggleClass("is-selected", !this.folderSel);
		const allIc = allRow.createSpan("pcal-folder-ic");
		setIcon(allIc, "inbox");
		allRow.createSpan({ cls: "pcal-folder-name", text: "All inboxes" });
		allRow.addEventListener("click", () => {
			this.plugin.clearMailSearch();
			this.folderSel = null;
			this.screen = "list";
			this.autoSelectPending = true;
			this.render();
		});
		const collapsed = new Set(this.plugin.settings.mailCollapsed);
		const toggleCollapse = (key: string) => {
			const next = new Set(this.plugin.settings.mailCollapsed);
			if (next.has(key)) next.delete(key);
			else next.add(key);
			this.plugin.settings.mailCollapsed = [...next];
			this.plugin.queueSave();
			this.render();
		};
		// pinned folders first, Outlook's Favorites: any folder from any
		// account, in the user's own order, reordered by dragging
		const favs = this.plugin.settings.mailFavorites.filter((f) => accounts.some((x) => x.id === f.accountId));
		if (favs.length) {
			host.createDiv({ cls: "pcal-folder-favhead", text: "Favorites" });
			// the account tag only earns its place when two favorites collide
			// on the same display name (two accounts' Unread Mail, say)
			const dispCounts = new Map<string, number>();
			for (const f of favs) {
				const ac = accounts.find((x) => x.id === f.accountId);
				if (!ac) continue;
				const rn = f.folderId === UNREAD_FOLDER ? "Unread Mail" : this.plugin.folderNamesFor(ac.id).get(f.folderId) ?? "...";
				const dn = f.name?.trim() || rn;
				dispCounts.set(dn, (dispCounts.get(dn) ?? 0) + 1);
			}
			favs.forEach((fav, idx) => {
				const acc = accounts.find((x) => x.id === fav.accountId);
				if (!acc) return;
				const isUnread = fav.folderId === UNREAD_FOLDER;
				const realName = isUnread ? "Unread Mail" : this.plugin.folderNamesFor(acc.id).get(fav.folderId) ?? "...";
				const name = fav.name?.trim() || realName;
				const count = isUnread ? this.plugin.unreadSubtreeCount(acc) : this.plugin.folderUnreadRollup(acc.id, fav.folderId);
				const row = host.createDiv("pcal-folder-row pcal-fav-row");
				if (fav.indent) row.style.paddingLeft = "24px";
				row.toggleClass("is-selected", this.folderSel?.accountId === acc.id && this.folderSel?.folderId === fav.folderId);
				const ic = row.createSpan("pcal-folder-ic");
				setIcon(ic, isUnread ? "mail-open" : fav.folderId === this.plugin.inboxIdFor(acc) ? "inbox" : "folder");
				row.createSpan({ cls: "pcal-folder-name", text: name });
				if ((dispCounts.get(name) ?? 0) > 1) row.createSpan({ cls: "pcal-fav-acct", text: this.plugin.nameOf(acc) });
				if (count > 0) row.createSpan({ cls: "pcal-folder-count", text: String(count) });
				row.addEventListener("click", () => {
					this.plugin.clearMailSearch();
					this.folderSel = { accountId: acc.id, folderId: fav.folderId, name };
					this.screen = "list";
					this.autoSelectPending = true;
					this.plugin.ensureFolderMail(acc.id, fav.folderId, false);
					this.render();
				});
				row.addEventListener("contextmenu", (e) => {
					e.preventDefault();
					const menu = new Menu();
					if (idx > 0 || fav.indent)
						menu.addItem((i) =>
							i.setTitle(fav.indent ? "Outdent" : "Indent under the folder above").onClick(() => {
								this.plugin.settings.mailFavorites = this.plugin.settings.mailFavorites.map((f) => (f.accountId === fav.accountId && f.folderId === fav.folderId ? { ...f, indent: !fav.indent || undefined } : f));
								this.plugin.queueSave();
								this.render();
							})
						);
					menu.addItem((i) =>
						i.setTitle("Rename favorite").onClick(() =>
							new PromptModal(this.app, "Rename favorite", [{ label: "Name", value: fav.name ?? realName }], ([v]) => {
								this.plugin.settings.mailFavorites = this.plugin.settings.mailFavorites.map((f) => (f.accountId === fav.accountId && f.folderId === fav.folderId ? { ...f, name: v.trim() || undefined } : f));
								this.plugin.queueSave();
								this.render();
							}).open()
						)
					);
					menu.addItem((i) => i.setTitle("Remove from Favorites").onClick(() => this.toggleFavorite(acc.id, fav.folderId)));
					menu.showAtMouseEvent(e);
				});
				row.draggable = true;
				row.addEventListener("dragstart", () => (this.favDragIdx = idx));
				row.addEventListener("dragover", (e) => {
					e.preventDefault();
					row.addClass("pcal-drop-target");
				});
				row.addEventListener("dragleave", () => row.removeClass("pcal-drop-target"));
				row.addEventListener("drop", (e) => {
					e.preventDefault();
					row.removeClass("pcal-drop-target");
					const from = this.favDragIdx;
					this.favDragIdx = null;
					if (from == null || from === idx) return;
					const next = [...this.plugin.settings.mailFavorites];
					const [moved] = next.splice(from, 1);
					next.splice(idx, 0, moved);
					this.plugin.settings.mailFavorites = next;
					this.plugin.queueSave();
					this.render();
				});
			});
		}

		for (const a of accounts) {
			// the whole mailbox folds at the account level, Outlook-style
			const acctKey = `acct:${a.id}`;
			const acctCollapsed = collapsed.has(acctKey);
			const head = host.createDiv("pcal-folder-account");
			const twist = head.createSpan("pcal-folder-twist");
			setIcon(twist, acctCollapsed ? "chevron-right" : "chevron-down");
			const dot = head.createSpan("pcal-status-dot");
			dot.style.backgroundColor = colorOf.get(a.id) ?? "var(--interactive-accent)";
			head.createSpan({ cls: "pcal-folder-acctname", text: this.plugin.nameOf(a) });
			if (acctCollapsed) {
				const n = this.plugin.unreadSubtreeCount(a);
				if (n > 0) head.createSpan({ cls: "pcal-folder-count", text: String(n) });
			}
			head.addEventListener("click", () => toggleCollapse(acctKey));
			if (acctCollapsed) continue;
			const tree = this.plugin.folderTreeFor(a);
			if (!tree.length) {
				host.createDiv({ cls: "pcal-folder-empty", text: "Loading folders..." });
				continue;
			}
			const inboxId = this.plugin.inboxIdFor(a);
			let hideBelow = -1; // a hidden folder swallows its whole subtree
			for (const { folder, depth, expandable } of tree) {
				if (hideBelow >= 0 && depth > hideBelow) continue;
				hideBelow = -1;
				if (this.plugin.settings.mailHiddenFolders.some((h) => h.accountId === a.id && h.folderId === folder.id)) {
					hideBelow = depth;
					continue;
				}
				const row = host.createDiv("pcal-folder-row");
				row.style.paddingLeft = `${6 + depth * 14}px`;
				row.toggleClass("is-selected", this.folderSel?.accountId === a.id && this.folderSel?.folderId === folder.id);
				const isCollapsed = collapsed.has(folder.id);
				const twist = row.createSpan("pcal-folder-twist");
				if (expandable) {
					setIcon(twist, isCollapsed ? "chevron-right" : "chevron-down");
					twist.setAttribute("aria-label", isCollapsed ? "Expand" : "Collapse");
					twist.addEventListener("click", (e) => {
						e.stopPropagation();
						toggleCollapse(folder.id);
					});
				}
				const ic = row.createSpan("pcal-folder-ic");
				setIcon(ic, folder.id === inboxId ? "inbox" : "folder");
				row.createSpan({ cls: "pcal-folder-name", text: folder.name });
				// a collapsed branch rolls its subtree's unread up onto itself
				const count = isCollapsed && expandable ? this.plugin.folderUnreadRollup(a.id, folder.id) : folder.unread;
				if (count > 0) row.createSpan({ cls: "pcal-folder-count", text: String(count) });
				row.addEventListener("click", () => {
					this.plugin.clearMailSearch();
					this.folderSel = { accountId: a.id, folderId: folder.id, name: folder.name };
					this.screen = "list";
					this.autoSelectPending = true;
					this.plugin.ensureFolderMail(a.id, folder.id, false);
					this.render();
				});
				row.addEventListener("contextmenu", (e) => {
					e.preventDefault();
					const menu = new Menu();
					menu.addItem((i) => i.setTitle(this.isFavorite(a.id, folder.id) ? "Remove from Favorites" : "Add to Favorites").onClick(() => this.toggleFavorite(a.id, folder.id)));
					menu.addItem((i) =>
						i.setTitle("Hide folder").onClick(() => {
							this.plugin.settings.mailHiddenFolders = [...this.plugin.settings.mailHiddenFolders, { accountId: a.id, folderId: folder.id }];
							this.plugin.queueSave();
							this.render();
						})
					);
					menu.showAtMouseEvent(e);
				});
			}

			// hidden folders come back through their own quiet row
			const hidden = this.plugin.settings.mailHiddenFolders.filter((h) => h.accountId === a.id);
			if (hidden.length) {
				const hr = host.createDiv("pcal-folder-row pcal-folder-hiddenrow");
				hr.style.paddingLeft = "6px";
				hr.createSpan("pcal-folder-twist");
				const hic = hr.createSpan("pcal-folder-ic");
				setIcon(hic, "eye-off");
				hr.createSpan({ cls: "pcal-folder-name", text: "Hidden folders" });
				hr.createSpan({ cls: "pcal-folder-count", text: String(hidden.length) });
				hr.addEventListener("click", (e) => {
					const names = this.plugin.folderNamesFor(a.id);
					const menu = new Menu();
					for (const h of hidden)
						menu.addItem((i) =>
							i.setTitle(`Show ${names.get(h.folderId) ?? "..."}`).onClick(() => {
								this.plugin.settings.mailHiddenFolders = this.plugin.settings.mailHiddenFolders.filter((x) => !(x.accountId === h.accountId && x.folderId === h.folderId));
								this.plugin.queueSave();
								this.render();
							})
						);
					menu.showAtMouseEvent(e);
				});
			}

			// Search Folders close out the account, Outlook-style: a collapsible
			// branch holding the virtual Unread view and any saved searches
			const searchKey = `search:${a.id}`;
			const searchCollapsed = collapsed.has(searchKey);
			const shead = host.createDiv("pcal-folder-row pcal-folder-searchroot");
			shead.style.paddingLeft = "6px"; // the tree's depth-0 inset
			const stw = shead.createSpan("pcal-folder-twist");
			setIcon(stw, searchCollapsed ? "chevron-right" : "chevron-down");
			const sic = shead.createSpan("pcal-folder-ic");
			setIcon(sic, "search");
			shead.createSpan({ cls: "pcal-folder-name", text: "Search Folders" });
			shead.addEventListener("click", () => toggleCollapse(searchKey));
			if (searchCollapsed) continue;
			const unreadRow = host.createDiv("pcal-folder-row pcal-folder-unread");
			unreadRow.style.paddingLeft = "20px"; // the tree's depth-1 inset
			unreadRow.toggleClass("is-selected", this.folderSel?.accountId === a.id && this.folderSel?.folderId === UNREAD_FOLDER);
			unreadRow.createSpan("pcal-folder-twist"); // the empty chevron slot every tree row keeps
			const uic = unreadRow.createSpan("pcal-folder-ic");
			setIcon(uic, "mail-open");
			unreadRow.createSpan({ cls: "pcal-folder-name", text: "Unread Mail" });
			const unreadTotal = this.plugin.unreadSubtreeCount(a);
			if (unreadTotal > 0) unreadRow.createSpan({ cls: "pcal-folder-count", text: String(unreadTotal) });
			unreadRow.addEventListener("click", () => {
				this.plugin.clearMailSearch();
				this.folderSel = { accountId: a.id, folderId: UNREAD_FOLDER, name: "Unread Mail" };
				this.screen = "list";
				this.autoSelectPending = true;
				this.plugin.ensureFolderMail(a.id, UNREAD_FOLDER, false);
				this.render();
			});
			unreadRow.addEventListener("contextmenu", (e) => {
				e.preventDefault();
				const menu = new Menu();
				menu.addItem((i) => i.setTitle(this.isFavorite(a.id, UNREAD_FOLDER) ? "Remove from Favorites" : "Add to Favorites").onClick(() => this.toggleFavorite(a.id, UNREAD_FOLDER)));
				menu.showAtMouseEvent(e);
			});
			for (const sf of this.plugin.settings.mailSearchFolders.filter((x) => x.accountId === a.id)) {
				const row = host.createDiv("pcal-folder-row");
				row.style.paddingLeft = "20px";
				row.createSpan("pcal-folder-twist");
				const ic = row.createSpan("pcal-folder-ic");
				setIcon(ic, "search");
				row.createSpan({ cls: "pcal-folder-name", text: sf.name });
				row.addEventListener("click", () => {
					this.screen = "list";
					if (this.searchInputEl) this.searchInputEl.value = sf.query;
					void this.plugin.runMailSearch(sf.query);
					this.render();
				});
				row.addEventListener("contextmenu", (e) => {
					e.preventDefault();
					const menu = new Menu();
					menu.addItem((i) => i.setTitle("Edit search folder").onClick(() => new SearchFolderModal(this.app, this.plugin, a, sf, () => this.render()).open()));
					menu.addItem((i) =>
						i.setTitle("Delete search folder").onClick(() => {
							this.plugin.settings.mailSearchFolders = this.plugin.settings.mailSearchFolders.filter((x) => x.id !== sf.id);
							this.plugin.queueSave();
							this.render();
						})
					);
					menu.showAtMouseEvent(e);
				});
			}
			const addSearch = host.createDiv("pcal-folder-row pcal-folder-addsearch");
			addSearch.style.paddingLeft = "20px";
			addSearch.createSpan("pcal-folder-twist");
			const asIc = addSearch.createSpan("pcal-folder-ic");
			setIcon(asIc, "plus");
			addSearch.createSpan({ cls: "pcal-folder-name", text: "New search folder" });
			addSearch.addEventListener("click", () => new SearchFolderModal(this.app, this.plugin, a, null, () => this.render()).open());
		}
	}

	private isFavorite(accountId: string, folderId: string): boolean {
		return this.plugin.settings.mailFavorites.some((f) => f.accountId === accountId && f.folderId === folderId);
	}

	private toggleFavorite(accountId: string, folderId: string) {
		const s = this.plugin.settings;
		s.mailFavorites = this.isFavorite(accountId, folderId) ? s.mailFavorites.filter((f) => !(f.accountId === accountId && f.folderId === folderId)) : [...s.mailFavorites, { accountId, folderId }];
		this.plugin.queueSave();
		this.render();
	}

	/** auto = the view chose the message (folder entry), not the user: the
	 *  body shows, but nothing gets marked read on arrival and the phone
	 *  drill stays on its list. */
	private async select(m: PCMail, auto = false) {
		const s = this.plugin.settings;
		const prev = this.selected;
		this.autoSelectPending = false;
		this.clearReadTimer();
		if (prev?.id !== m.id) {
			// "when the selection changes" marks the message being LEFT
			if (s.markRead === "change" && prev?.unread && !this.autoReadSuppressed()) void this.plugin.setMailRead(prev, true);
		}
		this.selected = m;
		this.selectedBody = null;
		this.selectedAtts = null;
		this.inlineCids = null;
		if (this.drill && !auto) this.screen = "read";
		if (m.unread && !this.autoReadSuppressed()) {
			if (s.markRead === "select" && !auto) void this.plugin.setMailRead(m, true);
			else if (s.markRead === "delay") {
				this.readTimer = window.setTimeout(() => {
					this.readTimer = null;
					const cur = this.plugin.allMail().find((x) => x.id === m.id);
					if (this.selected?.id === m.id && cur?.unread) void this.plugin.setMailRead(cur, true);
				}, Math.max(1, s.markReadSeconds) * 1000);
			}
		}
		this.render();
		const body = await this.plugin.readMailBody(m);
		if (this.selected?.id !== m.id) return;
		this.selectedBody = body;
		this.renderReading();
		if (m.hasAttachments || (body?.html ?? "").includes("cid:")) {
			const atts = await this.plugin.mailAttachments(m);
			if (this.selected?.id !== m.id) return;
			this.selectedAtts = atts;
			this.renderReading();
			// inline images arrive as cid: references; the list call cannot see
			// contentId (a subtype property), so each inline attachment's own
			// fetch supplies it alongside the bytes
			const html = body?.html ?? "";
			const inline = atts.filter((a) => a.isInline).slice(0, 8);
			if (inline.length && html.includes("cid:")) {
				const map = new Map<string, string>();
				for (const att of inline) {
					const raw = await this.plugin.mailAttachmentRaw(m, att.id, true);
					if (this.selected?.id !== m.id) return;
					if (raw?.contentId && html.includes(raw.contentId)) map.set(raw.contentId, `data:${raw.contentType};base64,${raw.contentBytes}`);
				}
				if (map.size) {
					this.inlineCids = map;
					this.renderReading();
				}
			}
		}
	}

	private async openAttachment(m: PCMail, att: MailAttachment, open: boolean) {
		const path = await this.plugin.saveMailAttachment(m, att);
		if (!path) return;
		new Notice(`Power Desk: saved ${path}.`);
		if (open) (this.app as unknown as { openWithDefaultApp?: (p: string) => void }).openWithDefaultApp?.(path);
	}

	private async selectAndReply(m: PCMail, all = false) {
		await this.select(m);
		new RichComposeModal(this.app, this.plugin, { mode: all ? "replyAll" : "reply", mail: m }).open();
	}

	private renderReading() {
		const host = this.readEl;
		if (!host) return;
		host.empty();
		const m = this.selected;
		for (const b of this.mailToolBtns) b.toggleClass("is-disabled", !m && this.multiSel.size === 0);
		if (!m) {
			host.createDiv({ cls: "pcal-embed-empty", text: "Select a message." });
			return;
		}
		host.createDiv({ cls: "pcal-mail-read-subject", text: m.subject });
		host.createDiv({ cls: "pcal-mail-read-meta", text: `${m.from}${m.fromAddress ? ` <${m.fromAddress}>` : ""}` });
		if (this.selectedBody?.toLine) host.createDiv({ cls: "pcal-mail-read-meta", text: `To: ${this.selectedBody.toLine}` });

		const btns = host.createDiv("pcal-card-btns");
		const reply = btns.createEl("button", { text: "Reply", cls: "mod-cta" });
		btns.createEl("button", { text: "Archive" }).addEventListener("click", () => {
			void this.plugin.archiveMail(m);
			this.selected = null;
			this.render();
		});
		btns.createEl("button", { text: m.unread ? "Mark read" : "Mark unread" }).addEventListener("click", () => void this.plugin.setMailRead(m, m.unread));
		btns.createEl("button", { text: "Make event" }).addEventListener("click", () => {
			const now = Date.now();
			const min = Math.ceil((minutesOfMs(now) + 1) / 30) * 30;
			const dayKey = min >= 1440 ? addDays(keyOfMs(now), 1) : keyOfMs(now);
			const start = msOfKey(dayKey) + (min % 1440) * 60000;
			new EventModal(this.app, this.plugin, null, start, start + 30 * 60000, false, { title: subjectToEventTitle(m.subject), invites: m.fromAddress }).open();
		});
		btns.createEl("button", { text: "Save to note" }).addEventListener("click", () => void this.plugin.saveMailToNote(m));
		if (m.webLink) btns.createEl("button", { text: "Open in Outlook" }).addEventListener("click", () => window.open(m.webLink, "_blank"));

		// attachments ride their own bar under the actions, Outlook-style
		const files = (this.selectedAtts ?? []).filter((a) => !a.isInline);
		if ((m.hasAttachments && !this.selectedAtts) || files.length) {
			const bar = host.createDiv("pcal-mail-atts");
			if (!this.selectedAtts) bar.createSpan({ cls: "pcal-mail-att-note", text: "Loading attachments..." });
			for (const att of files) {
				const chip = bar.createDiv("pcal-mail-att");
				const ic = chip.createSpan("pcal-mail-att-icon");
				setIcon(ic, att.contentType.startsWith("image/") ? "image" : att.contentType.includes("pdf") ? "file-text" : "paperclip");
				const tx = chip.createDiv("pcal-mail-att-text");
				tx.createDiv({ cls: "pcal-mail-att-name", text: att.name });
				tx.createDiv({ cls: "pcal-mail-att-size", text: att.size >= 1048576 ? `${(att.size / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(att.size / 1024))} KB` });
				chip.addEventListener("click", (e) => {
					const menu = new Menu();
					menu.addItem((i) => i.setTitle("Preview").onClick(() => void this.openAttachment(m, att, true)));
					menu.addItem((i) => i.setTitle("Save to vault").onClick(() => void this.openAttachment(m, att, false)));
					if (Platform.isDesktopApp) menu.addItem((i) => i.setTitle("Save to folder").onClick(() => void this.plugin.saveMailAttachmentLocal(m, att)));
					menu.showAtMouseEvent(e);
				});
			}
		}

		reply.addEventListener("click", () => new RichComposeModal(this.app, this.plugin, { mode: "reply", mail: m }).open());

		// HTML mail renders sanitized on its own light card; plain text stays text
		let html = this.selectedBody?.html ?? "";
		if (html && this.inlineCids) for (const [cid, url] of this.inlineCids) html = html.split(`cid:${cid}`).join(url);
		const bodyHost = host.createDiv("pcal-mail-read-body");
		if (html) bodyHost.createDiv("pcal-mail-html").appendChild(sanitizeHTMLToDom(html));
		else bodyHost.setText(this.selectedBody ? this.selectedBody.text : "Loading...");
	}
}

/** A small labeled-fields prompt: favorite renames, new search folders. */
class PromptModal extends Modal {
	constructor(
		app: App,
		private heading: string,
		private fields: { label: string; value: string; placeholder?: string }[],
		private onSubmit: (values: string[]) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText(this.heading);
		const inputs: HTMLInputElement[] = [];
		for (const f of this.fields)
			new Setting(this.contentEl).setName(f.label).addText((t) => {
				t.setValue(f.value);
				if (f.placeholder) t.setPlaceholder(f.placeholder);
				inputs.push(t.inputEl);
			});
		const btns = this.contentEl.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			this.onSubmit(inputs.map((i) => i.value));
			this.close();
		});
		window.setTimeout(() => inputs[0]?.focus(), 20);
	}

	onClose() {
		this.contentEl.empty();
	}
}

/* ---------------- the calendar view ---------------- */

const HOUR_H = 48;

class PowerCalendarView extends ItemView {
	private mode: ViewMode;
	private anchorKey: string;
	private titleEl!: HTMLElement;
	private bodyEl!: HTMLElement;
	private statusEl!: HTMLElement;
	private modeBtns = new Map<ViewMode, HTMLElement>();
	private refreshBtn!: HTMLElement;
	private newBtn!: HTMLElement;
	private filterBtn!: HTMLElement;
	private dayLabelEl: HTMLElement | null = null;
	private sidebarEl!: HTMLElement;
	private fullscreen = false;
	/** The sidebar mini-month's own month, navigable without moving the view. */
	private sbAnchor = keyOfDate(new Date()).slice(0, 7) + "-01";
	private sbRenderQueued = false;
	private cardEl: HTMLElement | null = null;
	private cardCleanup: (() => void) | null = null;
	private miniEl: HTMLElement | null = null;
	private miniCleanup: (() => void) | null = null;
	private miniAnchor = "";
	private renderQueued = false;
	private todayKey = keyOfDate(new Date());
	private readonly onData = () => this.queueRender();
	private readonly onMiniData = () => this.renderMiniGrid();

	constructor(
		leaf: WorkspaceLeaf,
		private plugin: PowerDeskPlugin
	) {
		super(leaf);
		this.mode = Platform.isPhone ? plugin.settings.phoneDefaultMode : plugin.settings.defaultMode;
		this.anchorKey = keyOfDate(new Date());
		this.scope = new Scope(this.app.scope);
		const key = (k: string, cb: () => void) =>
			(this.scope as Scope).register([], k, (evt) => {
				const t = evt.target as HTMLElement | null;
				if (t && (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t.isContentEditable)) return true;
				cb();
				return false;
			});
		key("T", () => this.goToday());
		key("M", () => this.setMode("month"));
		key("W", () => this.setMode("week"));
		key("5", () => this.setMode("workweek"));
		key("D", () => this.setMode("day"));
		key("A", () => this.setMode("agenda"));
		key("R", () => this.refresh(true));
		key("C", () => this.quickCreate());
		key("/", () => new EventFindModal(this.app, this.plugin).open());
		key("ArrowLeft", () => this.step(-1));
		key("ArrowRight", () => this.step(1));
		key("F", () => this.toggleFullscreen());
		key("S", () => this.toggleSidebar());
		(this.scope as Scope).register([], "Escape", () => {
			if (this.fullscreen && !this.cardEl && !this.miniEl) {
				this.toggleFullscreen();
				return false;
			}
			return true;
		});
	}

	getViewType(): string {
		return VIEW_TYPE;
	}

	getDisplayText(): string {
		return "Calendar";
	}

	getIcon(): string {
		return "calendar-days";
	}

	async onOpen() {
		this.plugin.listeners.add(this.onData);
		const root = this.contentEl;
		root.empty();
		root.addClass("pcal-root");

		const header = root.createDiv("pcal-header");
		if (!Platform.isPhone) {
			const sbBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Toggle the sidebar (S)" } });
			setIcon(sbBtn, "panel-left");
			sbBtn.addEventListener("click", () => this.toggleSidebar());
		}
		const nav = header.createDiv("pcal-nav");
		const prev = nav.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Previous (left arrow)" } });
		setIcon(prev, "chevron-left");
		prev.addEventListener("click", () => this.step(-1));
		const next = nav.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Next (right arrow)" } });
		setIcon(next, "chevron-right");
		next.addEventListener("click", () => this.step(1));
		const today = nav.createEl("button", { cls: "pcal-today-btn", text: "Today", attr: { "aria-label": "Go to today (T)" } });
		today.addEventListener("click", () => this.goToday());

		this.titleEl = header.createDiv("pcal-title");
		this.titleEl.setAttribute("aria-label", "Jump to a month");
		this.titleEl.addEventListener("click", () => this.openMini());

		const right = header.createDiv("pcal-header-right");
		this.newBtn = right.createEl("button", { cls: "pcal-new-btn", attr: { "aria-label": "New event (C); the arrow offers mail too" } });
		this.newBtn.createSpan({ text: "New event" });
		const newCaret = this.newBtn.createSpan("pcal-mode-caret");
		setIcon(newCaret, "chevron-down");
		this.newBtn.addEventListener("click", (e) => {
			if (e.target instanceof Node && newCaret.contains(e.target)) {
				const menu = new Menu();
				menu.addItem((i) => i.setTitle("Event").onClick(() => this.quickCreate()));
				menu.addItem((i) => i.setTitle("Mail").onClick(() => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open()));
				menu.showAtMouseEvent(e);
				return;
			}
			this.quickCreate();
		});
		const modes = right.createDiv("pcal-modes");
		const modeBtn = (m: ViewMode, label: string, hint: string, onClick?: (e: MouseEvent) => void) => {
			const b = modes.createEl("button", { cls: "pcal-mode-btn", text: label, attr: { "aria-label": `${label} view (${hint})` } });
			b.addEventListener("click", (e) => (onClick ? onClick(e) : this.setMode(m)));
			this.modeBtns.set(m, b);
		};
		modeBtn("month", "Month", "M");
		modeBtn("week", "Week", "W");
		modeBtn("workweek", "Work week", "5");
		// Day wears a visible dropdown arrow: the label switches the view, the
		// arrow picks how many days it spans, Outlook-style
		const dayBtn = modes.createEl("button", { cls: "pcal-mode-btn", attr: { "aria-label": "Day view (D); the arrow picks 1-7 days" } });
		this.dayLabelEl = dayBtn.createSpan({ text: "Day" });
		const dayCaret = dayBtn.createSpan("pcal-mode-caret");
		setIcon(dayCaret, "chevron-down");
		dayBtn.addEventListener("click", (e) => {
			if ((e.target instanceof Node && dayCaret.contains(e.target)) || this.mode === "day") this.openDayCountMenu(e);
			else this.setMode("day");
		});
		this.modeBtns.set("day", dayBtn);
		modeBtn("agenda", "Agenda", "A");
		this.filterBtn = right.createEl("button", { cls: "pcal-icon-btn pcal-filter-btn", attr: { "aria-label": "Filter events" } });
		const filterIc = this.filterBtn.createSpan();
		setIcon(filterIc, "filter");
		this.filterBtn.createSpan({ text: "Filter" });
		this.filterBtn.addEventListener("click", (e) => this.openFilterMenu(e));
		this.refreshBtn = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Refresh (R)" } });
		setIcon(this.refreshBtn, "refresh-cw");
		this.refreshBtn.addEventListener("click", () => this.refresh(true));
		const splitR = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Split right" } });
		setIcon(splitR, "separator-vertical");
		splitR.addEventListener("click", () => void this.app.workspace.duplicateLeaf(this.leaf, "split", "vertical"));
		const splitD = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Split down" } });
		setIcon(splitD, "separator-horizontal");
		splitD.addEventListener("click", () => void this.app.workspace.duplicateLeaf(this.leaf, "split", "horizontal"));
		const gear = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Power Desk settings" } });
		setIcon(gear, "settings");
		gear.addEventListener("click", () => {
			const st = (this.app as unknown as { setting?: { open: () => void; openTabById: (id: string) => void } }).setting;
			st?.open();
			st?.openTabById("powerdesk");
		});

		const main = root.createDiv("pcal-main");
		this.sidebarEl = main.createDiv("pcal-sidebar");
		this.sidebarEl.toggleClass("is-hidden", !this.plugin.settings.sidebarOpen || Platform.isPhone);
		// section headers toggle through delegation on the persistent rail, so
		// a rebuild racing the click can never eat it
		this.sidebarEl.addEventListener("click", (e) => {
			// Element, not HTMLElement: the chevron is an SVG and its paths are
			// SVGElements, which still need to find the header above them
			const head = e.target instanceof Element ? e.target.closest<HTMLElement>("[data-sec]") : null;
			if (!head) return;
			const s = this.plugin.settings;
			if (head.dataset.sec === "cals") s.sidebarCalsCollapsed = !s.sidebarCalsCollapsed;
			else s.sidebarAgendaCollapsed = !s.sidebarAgendaCollapsed;
			this.plugin.queueSave();
			this.renderSidebar();
		});
		this.bodyEl = main.createDiv("pcal-body");
		this.statusEl = root.createDiv("pcal-status");

		// the minute tick moves the now line; a date change moves "today" itself
		this.registerInterval(
			window.setInterval(() => {
				const nowKey = keyOfDate(new Date());
				if (nowKey !== this.todayKey) {
					this.todayKey = nowKey;
					this.queueRender();
				} else if (this.mode === "week" || this.mode === "workweek" || this.mode === "day") {
					this.positionNowLine();
				}
			}, 30_000)
		);

		this.refresh(false);
	}

	async onClose() {
		this.containerEl.removeClass("pcal-fullscreen");
		this.plugin.listeners.delete(this.onData);
		this.closeCard();
		this.closeMini();
	}

	/* ---------------- mini month navigator ---------------- */

	private closeMini() {
		this.miniCleanup?.();
		this.miniCleanup = null;
		this.plugin.listeners.delete(this.onMiniData);
		this.miniEl?.remove();
		this.miniEl = null;
	}

	/** A month-at-a-glance popover under the title: density dots per day,
	 *  month arrows, click a day to land there in the current view. */
	private openMini() {
		if (this.miniEl) {
			this.closeMini();
			return;
		}
		this.closeCard();
		this.miniAnchor = `${this.anchorKey.slice(0, 7)}-01`;
		const pop = document.body.createDiv("pcal-mini");
		this.miniEl = pop;
		const head = pop.createDiv("pcal-mini-head");
		const prev = head.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Previous month" } });
		setIcon(prev, "chevron-left");
		head.createDiv("pcal-mini-title");
		const next = head.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Next month" } });
		setIcon(next, "chevron-right");
		prev.addEventListener("click", () => {
			this.miniAnchor = stepAnchor("month", this.miniAnchor, -1);
			this.renderMiniGrid();
		});
		next.addEventListener("click", () => {
			this.miniAnchor = stepAnchor("month", this.miniAnchor, 1);
			this.renderMiniGrid();
		});
		pop.createDiv("pcal-mini-grid");
		const r = this.titleEl.getBoundingClientRect();
		pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 272)) + "px";
		pop.style.top = r.bottom + 6 + "px";
		this.plugin.listeners.add(this.onMiniData);
		const onDown = (e: MouseEvent) => {
			if (e.target instanceof Node && (pop.contains(e.target) || this.titleEl.contains(e.target))) return;
			this.closeMini();
		};
		const onKey = (e: KeyboardEvent) => {
			if (e.key === "Escape") this.closeMini();
		};
		document.addEventListener("pointerdown", onDown, true);
		document.addEventListener("keydown", onKey, true);
		this.miniCleanup = () => {
			document.removeEventListener("pointerdown", onDown, true);
			document.removeEventListener("keydown", onKey, true);
		};
		this.renderMiniGrid();
	}

	private renderMiniGrid() {
		const pop = this.miniEl;
		if (!pop || !pop.isConnected) return;
		const s = this.plugin.settings;
		const titleEl = pop.querySelector<HTMLElement>(".pcal-mini-title");
		const grid = pop.querySelector<HTMLElement>(".pcal-mini-grid");
		if (!titleEl || !grid) return;
		titleEl.setText(periodLabel("month", this.miniAnchor, s.weekStartsMonday));
		grid.empty();
		const cells = monthGrid(+this.miniAnchor.slice(0, 4), +this.miniAnchor.slice(5, 7) - 1, s.weekStartsMonday);
		this.plugin.ensureWindow(cells[0].key, cells[41].key, false);
		const events = this.plugin.eventsForWindow(cells[0].key, cells[41].key);
		for (const name of this.weekdayNames()) grid.createDiv({ cls: "pcal-mini-dow", text: name.slice(0, 1) });
		for (const cell of cells) {
			const d = grid.createDiv({ cls: "pcal-mini-day", text: String(cell.day) });
			d.toggleClass("is-other", !cell.inMonth);
			d.toggleClass("is-today", cell.key === this.todayKey);
			const n = eventsOnDay(events, cell.key).length;
			d.toggleClass("has-events", n > 0);
			d.toggleClass("has-many", n >= 4);
			d.addEventListener("click", () => {
				this.anchorKey = cell.key;
				this.closeMini();
				this.refresh(false);
			});
		}
	}

	goToday() {
		this.anchorKey = keyOfDate(new Date());
		this.refresh(false);
	}

	private toggleSidebar() {
		if (Platform.isPhone) return;
		const s = this.plugin.settings;
		s.sidebarOpen = !s.sidebarOpen;
		this.plugin.queueSave();
		this.sidebarEl.toggleClass("is-hidden", !s.sidebarOpen);
		if (s.sidebarOpen) this.renderSidebar();
	}

	/** The calendar over the whole window; F toggles, Esc leaves. */
	private toggleFullscreen() {
		this.fullscreen = !this.fullscreen;
		this.containerEl.toggleClass("pcal-fullscreen", this.fullscreen);
	}

	/** The left rail: a mini month for jumping, then every calendar with its
	 *  toggle. Click a row to toggle it; right-click offers solo and show-all. */
	private renderSidebar() {
		const host = this.sidebarEl;
		if (!host || host.hasClass("is-hidden")) return;
		host.empty();
		const s = this.plugin.settings;

		const mini = host.createDiv("pcal-sb-mini");
		const mh = mini.createDiv("pcal-mini-head");
		const prev = mh.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Previous month" } });
		setIcon(prev, "chevron-left");
		const mt = mh.createDiv({ cls: "pcal-mini-title" });
		const next = mh.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Next month" } });
		setIcon(next, "chevron-right");
		const grid = mini.createDiv("pcal-mini-grid");
		const paint = () => {
			mt.setText(periodLabel("month", this.sbAnchor, s.weekStartsMonday));
			grid.empty();
			const cells = monthGrid(+this.sbAnchor.slice(0, 4), +this.sbAnchor.slice(5, 7) - 1, s.weekStartsMonday);
			for (const name of this.weekdayNames()) grid.createDiv({ cls: "pcal-mini-dow", text: name.slice(0, 1) });
			for (const cell of cells) {
				const d = grid.createDiv({ cls: "pcal-mini-day", text: String(cell.day) });
				d.toggleClass("is-other", !cell.inMonth);
				d.toggleClass("is-today", cell.key === this.todayKey);
				d.toggleClass("is-anchor", cell.key === this.anchorKey);
				d.addEventListener("click", () => {
					this.anchorKey = cell.key;
					this.refresh(false);
				});
			}
		};
		prev.addEventListener("click", () => {
			this.sbAnchor = stepAnchor("month", this.sbAnchor, -1);
			paint();
		});
		next.addEventListener("click", () => {
			this.sbAnchor = stepAnchor("month", this.sbAnchor, 1);
			paint();
		});
		paint();

		// the calendar list folds behind its header, buying the agenda room
		const calsHead = host.createDiv({ cls: "pcal-sb-section", attr: { "data-sec": "cals" } });
		const chev = calsHead.createSpan("pcal-sb-chev");
		setIcon(chev, s.sidebarCalsCollapsed ? "chevron-right" : "chevron-down");
		calsHead.createSpan({ text: "Calendars" });
		for (const g of this.sidebarGroups()) {
			if (s.sidebarCalsCollapsed) break;
			host.createDiv({ cls: "pcal-sb-group", text: g.label });
			for (const it of g.items) {
				const row = host.createDiv({ cls: "pcal-sb-row", attr: { "aria-label": "Click to toggle; right-click for more" } });
				row.toggleClass("is-off", !it.enabled);
				row.createSpan("pcal-sb-dot").style.background = it.color || "var(--interactive-accent)";
				row.createSpan({ cls: "pcal-sb-name", text: it.name });
				row.addEventListener("click", () => {
					it.setEnabled(!it.enabled);
					this.afterSourceToggle();
				});
				row.addEventListener("contextmenu", (e) => {
					e.preventDefault();
					const menu = new Menu();
					menu.addItem((i) => {
						i.setTitle("Color");
						// setSubmenu shipped after the bundled typings; the runtime has it
						const sub = (i as unknown as { setSubmenu?: () => Menu }).setSubmenu?.();
						if (!sub) return;
						PALETTE.forEach((hex, idx) => {
							sub.addItem((si) => {
								const frag = document.createDocumentFragment();
								const dot = document.createElement("span");
								dot.className = "pcal-sb-menu-dot";
								dot.style.background = hex;
								frag.appendChild(dot);
								frag.appendChild(document.createTextNode(SIDEBAR_COLOR_NAMES[idx] ?? hex));
								si.setTitle(frag)
									.setChecked(it.color.toLowerCase() === hex)
									.onClick(() => {
										it.setColor(hex);
										this.afterSourceToggle();
									});
							});
						});
					});
					menu.addSeparator();
					menu.addItem((i) =>
						i.setTitle("Show only this calendar").onClick(() => {
							this.setAllCalendars(false);
							it.setEnabled(true);
							this.afterSourceToggle();
						})
					);
					menu.addItem((i) =>
						i.setTitle("Show all calendars").onClick(() => {
							this.setAllCalendars(true);
							this.afterSourceToggle();
						})
					);
					menu.showAtMouseEvent(e);
				});
			}
		}

		// the coming week at a glance under everything, Fantastical-style:
		// day headers wear the forecast, rows open the event card
		const agHead = host.createDiv({ cls: "pcal-sb-section pcal-sb-agenda-head", attr: { "data-sec": "agenda" } });
		const agChev = agHead.createSpan("pcal-sb-chev");
		setIcon(agChev, s.sidebarAgendaCollapsed ? "chevron-right" : "chevron-down");
		agHead.createSpan({ text: "Agenda" });
		if (s.sidebarAgendaCollapsed) return;
		const ag = host.createDiv("pcal-sb-agenda");
		const from = this.todayKey;
		this.plugin.ensureWindow(from, addDays(from, 6), false);
		const agEvents = this.plugin.eventsForWindow(from, addDays(from, 6));
		for (let i = 0; i < 7; i++) {
			const key = addDays(from, i);
			const dayEvents = eventsOnDay(agEvents, key);
			const head = ag.createDiv("pcal-sb-dayhead");
			head.createSpan({ cls: "pcal-sb-dayname", text: (i === 0 ? "Today" : i === 1 ? "Tomorrow" : DAYS_SHORT[dayOfWeek(key)]) + ` ${+key.slice(5, 7)}/${+key.slice(8, 10)}` });
			const w = this.plugin.weatherFor(key);
			if (w) head.createSpan({ cls: "pcal-sb-weather", text: `${weatherGlyph(w.code)} ${w.hi}°/${w.lo}°` });
			head.addEventListener("click", () => {
				this.anchorKey = key;
				this.refresh(false);
			});
			if (!dayEvents.length) {
				ag.createDiv({ cls: "pcal-sb-noevents", text: "No events" });
				continue;
			}
			for (const ev of dayEvents) {
				const row = ag.createDiv("pcal-sb-ev");
				row.createSpan("pcal-sb-dot").style.background = ev.color || "var(--interactive-accent)";
				const tx = row.createDiv("pcal-sb-ev-text");
				tx.createDiv({ cls: "pcal-sb-ev-title", text: ev.title });
				tx.createDiv({ cls: "pcal-sb-ev-time", text: ev.allDay ? "All day" : fmtEventRange(ev, s.use24h) });
				row.addEventListener("click", (e) => this.openCard(ev, e.currentTarget as HTMLElement));
			}
		}
	}

	private sidebarGroups(): { label: string; items: SidebarCal[] }[] {
		const s = this.plugin.settings;
		const groups: { label: string; items: SidebarCal[] }[] = [];
		for (const a of s.graphAccounts)
			if (a.calendars.length)
				groups.push({
					label: this.plugin.nameOf(a),
					items: a.calendars.map((c) => ({ name: c.name, color: c.color, enabled: c.enabled, setEnabled: (v: boolean) => (c.enabled = v), setColor: (v: string) => (c.color = v) })),
				});
		for (const g of s.googleAccounts)
			if (g.calendars.length)
				groups.push({
					label: this.plugin.nameOf(g),
					items: g.calendars.map((c) => ({ name: c.name, color: c.color, enabled: c.enabled, setEnabled: (v: boolean) => (c.enabled = v), setColor: (v: string) => (c.color = v) })),
				});
		for (const cd of s.caldavAccounts)
			if (cd.collections.length)
				groups.push({
					label: cd.name,
					items: cd.collections.map((c) => ({ name: c.name, color: c.color, enabled: c.enabled, setEnabled: (v: boolean) => (c.enabled = v), setColor: (v: string) => (c.color = v) })),
				});
		if (s.icsFeeds.length)
			groups.push({ label: "ICS feeds", items: s.icsFeeds.map((f) => ({ name: f.name, color: f.color, enabled: f.enabled, setEnabled: (v: boolean) => (f.enabled = v), setColor: (v: string) => (f.color = v) })) });
		if (s.vaultSources.length)
			groups.push({ label: "Vault notes", items: s.vaultSources.map((v) => ({ name: v.name, color: v.color, enabled: v.enabled, setEnabled: (x: boolean) => (v.enabled = x), setColor: (x: string) => (v.color = x) })) });
		return groups;
	}

	private setAllCalendars(v: boolean) {
		for (const g of this.sidebarGroups()) for (const it of g.items) it.setEnabled(v);
	}

	private afterSourceToggle() {
		this.plugin.queueSave();
		this.plugin.sourcesChanged();
		this.renderSidebar();
		this.queueRender();
	}

	setMode(mode: ViewMode) {
		if (this.mode === mode) return;
		this.mode = mode;
		this.refresh(false);
	}

	/** Jump straight to one day (a month cell's number, a week header). */
	goDay(key: string) {
		this.anchorKey = key;
		this.mode = "day";
		this.refresh(false);
	}

	private step(dir: 1 | -1) {
		this.anchorKey = stepAnchor(this.mode, this.anchorKey, dir, this.plugin.settings.agendaDays, this.plugin.settings.dayViewDays);
		this.refresh(false);
	}

	private openDayCountMenu(e: MouseEvent) {
		const menu = new Menu();
		for (let n = 1; n <= 7; n++)
			menu.addItem((i) =>
				i
					.setTitle(n === 1 ? "1 day" : `${n} days`)
					.setChecked(this.plugin.settings.dayViewDays === n)
					.onClick(() => {
						this.plugin.settings.dayViewDays = n;
						this.mode = "day";
						this.plugin.queueSave();
						this.refresh(false);
					})
			);
		menu.showAtMouseEvent(e);
	}

	/** No filter applied means every box is checked; anything else lights the
	 *  funnel button so hidden events are never a mystery. */
	private filterActive(): boolean {
		const s = this.plugin.settings;
		return !(s.filterMeetings && s.filterAppointments && s.filterAllDay && s.filterOnline && s.filterTentative) || s.filterHiddenCategories.length > 0;
	}

	private applyFilter(events: PCEvent[]): PCEvent[] {
		if (!this.filterActive()) return events;
		const s = this.plugin.settings;
		return events.filter((ev) => {
			const meeting = (ev.attendeeDetail?.length ?? ev.attendees?.length ?? 0) > 0;
			if (!s.filterAllDay && ev.allDay) return false;
			if (!s.filterOnline && ev.joinUrl) return false;
			if (!s.filterTentative && ev.tentative) return false;
			if (!s.filterMeetings && meeting && !ev.allDay) return false;
			if (!s.filterAppointments && !meeting && !ev.allDay) return false;
			if (s.filterHiddenCategories.length && ev.categories?.some((c) => s.filterHiddenCategories.includes(c))) return false;
			return true;
		});
	}

	private openFilterMenu(e: MouseEvent) {
		const s = this.plugin.settings;
		const menu = new Menu();
		const row = (title: string, get: () => boolean, set: (v: boolean) => void) =>
			menu.addItem((i) =>
				i
					.setTitle(title)
					.setChecked(get())
					.onClick(() => {
						set(!get());
						this.plugin.queueSave();
						this.queueRender();
					})
			);
		row("Meetings", () => s.filterMeetings, (v) => (s.filterMeetings = v));
		row("Appointments", () => s.filterAppointments, (v) => (s.filterAppointments = v));
		row("All-day events", () => s.filterAllDay, (v) => (s.filterAllDay = v));
		row("Online meetings", () => s.filterOnline, (v) => (s.filterOnline = v));
		row("Tentative", () => s.filterTentative, (v) => (s.filterTentative = v));
		// every category seen in the loaded window, plus any already hidden so
		// they stay reachable to turn back on
		const win = viewWindow(this.mode, this.anchorKey, s.weekStartsMonday, s.agendaDays, s.dayViewDays);
		const cats = new Set<string>(s.filterHiddenCategories);
		for (const ev of this.plugin.eventsForWindow(win.fromKey, win.toKey)) for (const c of ev.categories ?? []) cats.add(c);
		if (cats.size)
			menu.addItem((i) => {
				i.setTitle("Categories");
				const sub = (i as unknown as { setSubmenu?: () => Menu }).setSubmenu?.();
				if (!sub) return;
				for (const c of [...cats].sort())
					sub.addItem((si) =>
						si
							.setTitle(c)
							.setChecked(!s.filterHiddenCategories.includes(c))
							.onClick(() => {
								s.filterHiddenCategories = s.filterHiddenCategories.includes(c) ? s.filterHiddenCategories.filter((x) => x !== c) : [...s.filterHiddenCategories, c];
								this.plugin.queueSave();
								this.queueRender();
							})
					);
			});
		menu.addSeparator();
		menu.addItem((i) =>
			i
				.setTitle("Clear filters")
				.setDisabled(!this.filterActive())
				.onClick(() => {
					s.filterMeetings = s.filterAppointments = s.filterAllDay = s.filterOnline = s.filterTentative = true;
					s.filterHiddenCategories = [];
					this.plugin.queueSave();
					this.queueRender();
				})
		);
		menu.showAtMouseEvent(e);
	}

	/** New event at the next half hour (the C key and the + button). */
	private quickCreate() {
		if (!this.plugin.anyWritable()) {
			new Notice("Power Desk: connect an account with edit access in settings to create events.");
			return;
		}
		const now = Date.now();
		const min = Math.ceil((minutesOfMs(now) + 1) / 30) * 30;
		const dayKey = min >= 1440 ? addDays(keyOfMs(now), 1) : keyOfMs(now);
		const start = msOfKey(dayKey) + (min % 1440) * 60000;
		this.openEventModal(null, start, start + 30 * 60000, false);
	}

	private openEventModal(ev: PCEvent | null, startMs: number, endMs: number, allDay: boolean, prefill?: { title?: string; location?: string; invites?: string }) {
		new EventModal(this.app, this.plugin, ev, startMs, endMs, allDay, prefill).open();
	}

	refresh(force: boolean) {
		const s = this.plugin.settings;
		const win = viewWindow(this.mode, this.anchorKey, s.weekStartsMonday, s.agendaDays);
		this.plugin.ensureWindow(win.fromKey, win.toKey, force);
		this.render();
	}

	private queueRender() {
		if (this.renderQueued) return;
		this.renderQueued = true;
		requestAnimationFrame(() => {
			this.renderQueued = false;
			this.render();
		});
	}

	/* ---------------- rendering ---------------- */

	private render() {
		if (!this.bodyEl || !this.bodyEl.isConnected) return; // a queued frame can outlive the view
		const s = this.plugin.settings;
		this.closeCard();
		this.titleEl.setText(periodLabel(this.mode, this.anchorKey, s.weekStartsMonday, s.agendaDays, s.dayViewDays));
		for (const [m, b] of this.modeBtns) b.toggleClass("is-active", m === this.mode);
		this.dayLabelEl?.setText(s.dayViewDays > 1 ? `${s.dayViewDays} days` : "Day");
		this.filterBtn.toggleClass("is-active", this.filterActive());
		this.filterBtn.setAttribute("aria-label", this.filterActive() ? "Filter events (a filter is applied)" : "Filter events");
		this.refreshBtn.toggleClass("is-loading", this.plugin.anyInFlight());
		this.newBtn.toggle(this.plugin.anyWritable());

		const win = viewWindow(this.mode, this.anchorKey, s.weekStartsMonday, s.agendaDays, s.dayViewDays);
		// idempotent: covered-and-fresh sources are skipped, so a render triggered
		// by a fetch landing cannot start another fetch. This is what lets an
		// externally-synced settings change (cache cleared) repopulate itself.
		this.plugin.ensureWindow(win.fromKey, win.toKey, false);
		const events = this.applyFilter(this.plugin.eventsForWindow(win.fromKey, win.toKey));

		this.bodyEl.empty();
		if (this.mode === "month") this.renderMonth(events);
		else if (this.mode === "agenda") this.renderAgenda(events, win.fromKey, win.toKey);
		else if (this.mode === "day") this.renderWeek(events, daySpanKeys(this.anchorKey, s.dayViewDays));
		else if (this.mode === "workweek") this.renderWeek(events, workWeekDays(this.anchorKey));
		else this.renderWeek(events, weekDays(this.anchorKey, s.weekStartsMonday));

		this.renderStatus();
		this.queueSidebar();
	}

	/** Data lands in bursts while sources fetch; the sidebar rebuilds once the
	 *  burst settles instead of churning under the pointer. */
	private queueSidebar() {
		if (this.sbRenderQueued) return;
		this.sbRenderQueued = true;
		window.setTimeout(() => {
			this.sbRenderQueued = false;
			if (this.bodyEl?.isConnected) this.renderSidebar();
		}, 250);
	}

	private weekdayNames(): string[] {
		const s = this.plugin.settings;
		return s.weekStartsMonday ? [...DAYS_SHORT.slice(1), DAYS_SHORT[0]] : [...DAYS_SHORT];
	}

	private renderMonth(events: PCEvent[]) {
		const s = this.plugin.settings;
		const cells = monthGrid(+this.anchorKey.slice(0, 4), +this.anchorKey.slice(5, 7) - 1, s.weekStartsMonday);
		const root = this.bodyEl.createDiv("pcal-month");
		const head = root.createDiv("pcal-month-head");
		for (const name of this.weekdayNames()) head.createDiv({ cls: "pcal-month-headcell", text: name });
		const grid = root.createDiv("pcal-month-grid");
		const cellMap = new Map<string, HTMLElement>();
		for (let r = 0; r < 6; r++) {
			const rowCells = cells.slice(r * 7, r * 7 + 7);
			this.renderMonthRow(grid, rowCells, events, cellMap);
		}
		this.attachMonthDragCreate(grid, cellMap);
	}

	/** Drag across month cells to draft a multi-day all-day event; a plain
	 *  click stays inert and double-click still drafts a single day. */
	private attachMonthDragCreate(grid: HTMLElement, cellMap: Map<string, HTMLElement>) {
		grid.addEventListener("pointerdown", (pd: PointerEvent) => {
			if (pd.button !== 0 || pd.pointerType !== "mouse" || !this.plugin.anyWritable()) return;
			const t = pd.target as HTMLElement;
			if (t.closest(".pcal-chip, .pcal-span, .pcal-month-daynum, button")) return;
			const startKey = t.closest<HTMLElement>(".pcal-month-cell")?.dataset.key;
			if (!startKey) return;
			let dragging = false;
			let b = startKey;
			const range = (): [string, string] => (dayDiff(startKey, b) >= 0 ? [startKey, b] : [b, startKey]);
			const paint = () => {
				const [lo, hi] = range();
				for (const [key, el] of cellMap) el.toggleClass("pcal-range", dayDiff(lo, key) >= 0 && dayDiff(key, hi) >= 0);
			};
			const move = (mv: PointerEvent) => {
				if (!dragging) {
					if (Math.abs(mv.clientX - pd.clientX) + Math.abs(mv.clientY - pd.clientY) < 6) return;
					dragging = true;
				}
				const over = (document.elementFromPoint(mv.clientX, mv.clientY) as HTMLElement | null)?.closest<HTMLElement>(".pcal-month-cell");
				if (over?.dataset.key) b = over.dataset.key;
				paint();
			};
			const up = () => {
				document.removeEventListener("pointermove", move);
				document.removeEventListener("pointerup", up);
				for (const el of cellMap.values()) el.removeClass("pcal-range");
				if (!dragging) return;
				const [lo, hi] = range();
				this.openEventModal(null, msOfKey(lo), msOfKey(addDays(hi, 1)), true);
			};
			document.addEventListener("pointermove", move);
			document.addEventListener("pointerup", up);
		});
	}

	private renderMonthRow(grid: HTMLElement, rowCells: DayCell[], events: PCEvent[], cellMap: Map<string, HTMLElement>) {
		const rowKeys = rowCells.map((c) => c.key);
		const spans = spansForRow(events, rowKeys);
		const laneCount = spans.reduce((m, sp) => Math.max(m, sp.lane + 1), 0);
		const row = grid.createDiv("pcal-month-row");
		if (this.plugin.settings.showWeekNumbers) row.createDiv({ cls: "pcal-weeknum", text: `W${isoWeekNum(rowKeys[0])}` });

		const cellsEl = row.createDiv("pcal-month-cells");
		for (const cell of rowCells) {
			const el = cellsEl.createDiv("pcal-month-cell");
			el.dataset.key = cell.key;
			cellMap.set(cell.key, el);
			el.toggleClass("is-other-month", !cell.inMonth);
			el.toggleClass("is-today", cell.key === this.todayKey);
			const num = el.createDiv({ cls: "pcal-month-daynum", text: String(cell.day) });
			num.addEventListener("click", () => this.goDay(cell.key));
			const chipArea = el.createDiv("pcal-month-chips");
			chipArea.style.marginTop = `${laneCount * 22}px`;
			el.addEventListener("dblclick", (e) => {
				if (e.target !== el && e.target !== chipArea) return;
				if (!this.plugin.anyWritable()) return;
				this.openEventModal(null, msOfKey(cell.key), msOfKey(addDays(cell.key, 1)), true);
			});
			const timed = timedOnDay(events, cell.key);
			const cap = 4;
			for (const ev of timed.slice(0, cap)) this.renderChip(chipArea, ev);
			if (timed.length > cap) {
				const more = chipArea.createEl("button", { cls: "pcal-more-btn", text: `+${timed.length - cap} more` });
				more.addEventListener("click", () => this.goDay(cell.key));
			}
		}

		if (spans.length) {
			const overlay = row.createDiv("pcal-row-spans");
			for (const sp of spans) {
				const el = overlay.createDiv("pcal-span");
				el.style.left = `calc(${(sp.startIdx / 7) * 100}% + 2px)`;
				el.style.width = `calc(${((sp.endIdx - sp.startIdx + 1) / 7) * 100}% - 4px)`;
				el.style.top = `${sp.lane * 22}px`;
				this.paintEventEl(el, sp.ev);
				el.toggleClass("continues-left", !sp.startsHere);
				el.toggleClass("continues-right", !sp.endsHere);
				el.createSpan({ cls: "pcal-span-title", text: sp.ev.title });
				if (this.plugin.noteExistsFor(sp.ev)) el.addClass("has-note");
				el.addEventListener("click", (e) => this.openCard(sp.ev, e.currentTarget as HTMLElement));
			}
		}
	}

	private renderChip(parent: HTMLElement, ev: PCEvent) {
		const s = this.plugin.settings;
		const chip = parent.createDiv("pcal-chip");
		chip.style.setProperty("--pcal-ev-color", ev.color ?? "var(--interactive-accent)");
		chip.toggleClass("is-tentative", !!ev.tentative);
		chip.toggleClass("is-declined", !!ev.declined);
		this.paintNeedsAction(chip, ev);
		chip.toggleClass("has-note", this.plugin.noteExistsFor(ev));
		chip.createSpan({ cls: "pcal-chip-time", text: fmtTimeOfMs(ev.startMs, s.use24h, true) });
		chip.createSpan({ cls: "pcal-chip-title", text: ev.title });
		chip.addEventListener("click", (e) => this.openCard(ev, e.currentTarget as HTMLElement));
	}

	private paintEventEl(el: HTMLElement, ev: PCEvent) {
		el.style.setProperty("--pcal-ev-color", ev.color ?? "var(--interactive-accent)");
		el.toggleClass("is-tentative", !!ev.tentative);
		el.toggleClass("is-declined", !!ev.declined);
		this.paintNeedsAction(el, ev);
	}

	/** An invite still waiting on an answer wears the awaiting-response wash. */
	private paintNeedsAction(el: HTMLElement, ev: PCEvent) {
		const needs = !!ev.canRsvp && ev.myResponse === "none";
		el.toggleClass("is-needsaction", needs);
		if (needs) el.style.setProperty("--pcal-needsaction", this.plugin.settings.calNeedsActionColor);
	}

	private renderWeek(events: PCEvent[], days: string[]) {
		const s = this.plugin.settings;
		const root = this.bodyEl.createDiv("pcal-week");
		root.toggleClass("pcal-single-day", days.length === 1);

		const head = root.createDiv("pcal-week-head");
		const spacer = head.createDiv("pcal-gutter-spacer");
		if (s.showWeekNumbers && days.length > 1) {
			spacer.addClass("pcal-weeknum-head");
			spacer.setText(`W${isoWeekNum(days[0])}`);
		}
		const headCells = head.createDiv("pcal-week-headcells");
		for (const key of days) {
			const cell = headCells.createDiv("pcal-week-headcell");
			cell.toggleClass("is-today", key === this.todayKey);
			cell.createSpan({ cls: "pcal-week-headname", text: DAYS_SHORT[dayOfWeek(key)] });
			cell.createSpan({ cls: "pcal-week-headnum", text: String(+key.slice(8, 10)) });
			cell.addEventListener("click", () => this.goDay(key));
		}

		// the all-day banner strip
		const spans = spansForRow(events, days);
		if (spans.length) {
			const strip = root.createDiv("pcal-allday");
			strip.createDiv({ cls: "pcal-gutter-spacer pcal-allday-label", text: "all-day" });
			const lanesEl = strip.createDiv("pcal-allday-lanes");
			const laneCount = spans.reduce((m, sp) => Math.max(m, sp.lane + 1), 0);
			lanesEl.style.height = `${laneCount * 24 + 2}px`;
			for (const sp of spans) {
				const el = lanesEl.createDiv("pcal-span");
				el.style.left = `calc(${(sp.startIdx / days.length) * 100}% + 2px)`;
				el.style.width = `calc(${((sp.endIdx - sp.startIdx + 1) / days.length) * 100}% - 4px)`;
				el.style.top = `${sp.lane * 24}px`;
				this.paintEventEl(el, sp.ev);
				el.toggleClass("continues-left", !sp.startsHere);
				el.toggleClass("continues-right", !sp.endsHere);
				el.createSpan({ cls: "pcal-span-title", text: sp.ev.title });
				if (this.plugin.noteExistsFor(sp.ev)) el.addClass("has-note");
				el.addEventListener("click", (e) => this.openCard(sp.ev, e.currentTarget as HTMLElement));
			}
		}

		const scroll = root.createDiv("pcal-week-scroll");
		const gutter = scroll.createDiv("pcal-gutter");
		const tz2 = s.secondTimeZone.trim();
		const dayStartMs = msOfKey(days[0]);
		const tz2Works = !!tz2 && fmtZoneClock(dayStartMs, tz2, s.use24h) != null;
		if (tz2Works) root.addClass("pcal-has-tz2");
		for (let h = 0; h < 24; h++) {
			const lab = gutter.createDiv("pcal-gutter-hour");
			lab.style.top = `${h * HOUR_H}px`;
			if (h > 0) {
				if (tz2Works) lab.createSpan({ cls: "pcal-hour-tz2", text: fmtZoneClock(dayStartMs + h * 3600000, tz2, s.use24h) ?? "" });
				lab.createSpan({ cls: "pcal-hour-local", text: fmtClock(h * 60, s.use24h, true) });
			}
		}
		const cols = scroll.createDiv("pcal-week-cols");
		const colEls: HTMLElement[] = [];
		for (const key of days) {
			const col = cols.createDiv("pcal-week-col");
			colEls.push(col);
			col.toggleClass("is-today", key === this.todayKey);
			this.attachSlotGesture(col, key);
			const items = eventsOnDay(events, key)
				.filter((ev) => !isSpanEvent(ev))
				.map((ev) => ({ ev, c: clipToDay(ev, key) }))
				.filter((x): x is { ev: PCEvent; c: { startMin: number; endMin: number } } => x.c != null);
			const packed = packColumns(items.map((x) => x.c));
			items.forEach(({ ev, c }, i) => {
				const p = packed[i];
				const block = col.createDiv("pcal-block");
				block.style.top = `${(c.startMin / 60) * HOUR_H}px`;
				block.style.height = `${Math.max(18, ((c.endMin - c.startMin) / 60) * HOUR_H - 2)}px`;
				block.style.left = `calc(${(p.col / p.cols) * 100}% + 2px)`;
				block.style.width = `calc(${(1 / p.cols) * 100}% - 5px)`;
				this.paintEventEl(block, ev);
				block.toggleClass("has-note", this.plugin.noteExistsFor(ev));
				block.createDiv({ cls: "pcal-block-title", text: ev.title });
				if (c.endMin - c.startMin >= 40) block.createDiv({ cls: "pcal-block-time", text: fmtEventRange(ev, s.use24h) });
				this.attachBlockGesture(block, ev, key, days, colEls);
			});
			if (key === this.todayKey) {
				const line = col.createDiv("pcal-now-line");
				line.createDiv("pcal-now-dot");
			}
		}
		// the clock ruler rides every timed view: a faint line across the columns
		// and the time in the gutter, wherever the view is anchored. Only
		// today's own column carries the bold line above. A work week viewed on
		// a Saturday would otherwise show no timeline at all.
		cols.createDiv("pcal-now-line-week");
		gutter.createDiv("pcal-now-badge");
		this.positionNowLine();
		requestAnimationFrame(() => {
			scroll.scrollTop = Math.max(0, this.plugin.settings.dayStartHour * HOUR_H - 8);
		});
	}

	private positionNowLine() {
		const mins = minutesOfMs(Date.now());
		const top = `${(mins / 60) * HOUR_H}px`;
		for (const line of Array.from(this.bodyEl.querySelectorAll<HTMLElement>(".pcal-now-line, .pcal-now-line-week"))) {
			line.style.top = top;
		}
		const badge = this.bodyEl.querySelector<HTMLElement>(".pcal-now-badge");
		if (badge) {
			badge.style.top = top;
			badge.setText(fmtTimeOfMs(Date.now(), this.plugin.settings.use24h, true));
		}
	}

	private renderAgenda(events: PCEvent[], fromKey: string, toKey: string) {
		const s = this.plugin.settings;
		const root = this.bodyEl.createDiv("pcal-agenda");
		const groups = groupByDay(events, fromKey, toKey);
		if (!groups.length) {
			const empty = root.createDiv("pcal-empty");
			if (!this.plugin.sources().length) {
				empty.createDiv({ text: "No calendar sources are set up yet." });
				const b = empty.createEl("button", { text: "Open settings", cls: "mod-cta" });
				b.addEventListener("click", () => this.plugin.openOwnSettings());
			} else {
				empty.createDiv({ text: "Nothing scheduled in this window." });
			}
			return;
		}
		for (const g of groups) {
			const day = root.createDiv("pcal-agenda-day");
			day.toggleClass("is-today", g.key === this.todayKey);
			const head = day.createDiv({ cls: "pcal-agenda-head", text: g.key === this.todayKey ? `Today, ${fmtDayHeading(g.key).split(", ")[1]}` : fmtDayHeading(g.key) });
			head.addEventListener("click", () => this.goDay(g.key));
			for (const ev of g.events) {
				const row = day.createDiv("pcal-agenda-row");
				this.paintEventEl(row, ev);
				row.toggleClass("has-note", this.plugin.noteExistsFor(ev));
				row.createDiv("pcal-agenda-dot");
				row.createDiv({ cls: "pcal-agenda-time", text: ev.allDay ? "All day" : fmtTimeOfMs(ev.startMs, s.use24h) });
				const main = row.createDiv("pcal-agenda-main");
				main.createDiv({ cls: "pcal-agenda-title", text: ev.title });
				const sub: string[] = [];
				if (ev.calendarName) sub.push(ev.calendarName);
				if (ev.location) sub.push(ev.location);
				if (sub.length) main.createDiv({ cls: "pcal-agenda-sub", text: sub.join(" · ") });
				if (ev.joinUrl) {
					const join = row.createEl("button", { cls: "pcal-icon-btn pcal-join-btn", attr: { "aria-label": "Join meeting" } });
					setIcon(join, "video");
					join.addEventListener("click", (e) => {
						e.stopPropagation();
						window.open(ev.joinUrl, "_blank");
					});
				}
				row.addEventListener("click", (e) => this.openCard(ev, e.currentTarget as HTMLElement));
			}
		}
	}

	private renderStatus() {
		this.statusEl.empty();
		const states = this.plugin.sourceStates();
		if (!states.length) {
			this.statusEl.createSpan({ cls: "pcal-status-hint", text: "No calendar sources connected." });
			const b = this.statusEl.createEl("button", { cls: "pcal-status-btn", text: "Set up sources" });
			b.addEventListener("click", () => this.plugin.openOwnSettings());
			return;
		}
		const chips = this.statusEl.createDiv("pcal-status-chips");
		let latest = 0;
		for (const { def, st } of states) {
			const chip = chips.createDiv("pcal-status-chip");
			const dot = chip.createSpan("pcal-status-dot");
			dot.style.backgroundColor = def.color;
			chip.createSpan({ text: def.label });
			if (st?.error) {
				chip.addClass("has-error");
				chip.setAttribute("aria-label", st.error);
			}
			if (st?.inFlight) chip.addClass("is-loading");
			if (st) latest = Math.max(latest, st.fetchedAt);
		}
		const right = this.statusEl.createSpan("pcal-status-updated");
		if (this.plugin.anyInFlight()) right.setText("Refreshing...");
		else if (latest) right.setText(`Updated ${fmtTimeOfMs(latest, this.plugin.settings.use24h)}`);
	}

	/* ---------------- write gestures (mouse; touch taps still open the card) ---------------- */

	/** Click opens the card; a mouse drag on an editable block moves it (or
	 *  resizes from the bottom edge), pointer-based like every Power drag
	 *  engine. Escape cancels; the block itself is the ghost, since the render
	 *  after commit or cancel rebuilds the truth anyway. */
	private attachBlockGesture(block: HTMLElement, ev: PCEvent, dayKey: string, days: string[], colEls: HTMLElement[]) {
		let suppressClick = false;
		block.addEventListener(
			"click",
			(e) => {
				if (suppressClick) {
					suppressClick = false;
					e.stopImmediatePropagation();
					e.preventDefault();
					return;
				}
				this.openCard(ev, block);
			},
			{ capture: true }
		);
		const editable = !!ev.canEdit && !ev.allDay;
		if (!editable) return;
		block.addClass("pcal-can-edit");
		block.createDiv("pcal-grip");
		block.addEventListener("pointerdown", (pd: PointerEvent) => {
			if (pd.button !== 0 || pd.pointerType !== "mouse") return;
			pd.preventDefault();
			const r0 = block.getBoundingClientRect();
			const mode: "move" | "resize" = pd.clientY > r0.bottom - 8 ? "resize" : "move";
			const startX = pd.clientX;
			const startY = pd.clientY;
			let dragging = false;
			let times = { startMs: ev.startMs, endMs: ev.endMs };
			let curIdx = Math.max(0, days.indexOf(dayKey));
			const timeEl = block.querySelector<HTMLElement>(".pcal-block-time");
			const move = (mv: PointerEvent) => {
				if (!dragging) {
					if (Math.abs(mv.clientY - startY) + Math.abs(mv.clientX - startX) < 5) return;
					dragging = true;
					block.addClass("pcal-dragging");
					this.closeCard();
				}
				const deltaMin = ((mv.clientY - startY) / HOUR_H) * 60;
				let idx = curIdx;
				if (mode === "move" && days.length > 1) {
					const hit = colEls.findIndex((c) => {
						const r = c.getBoundingClientRect();
						return mv.clientX >= r.left && mv.clientX < r.right;
					});
					if (hit >= 0) idx = hit;
				}
				times = dragTimes(ev.startMs, ev.endMs, days[idx], deltaMin, mode);
				if (idx !== curIdx && colEls[idx]) {
					colEls[idx].appendChild(block);
					curIdx = idx;
				}
				const clip = clipToDay(times, days[curIdx]);
				if (clip) {
					block.style.top = `${(clip.startMin / 60) * HOUR_H}px`;
					block.style.height = `${Math.max(18, ((clip.endMin - clip.startMin) / 60) * HOUR_H - 2)}px`;
				}
				timeEl?.setText(fmtEventRange({ ...ev, startMs: times.startMs, endMs: times.endMs }, this.plugin.settings.use24h));
			};
			const finish = (commit: boolean) => {
				document.removeEventListener("pointermove", move);
				document.removeEventListener("pointerup", up);
				document.removeEventListener("keydown", onKey, true);
				block.removeClass("pcal-dragging");
				if (!dragging) return; // the click handler takes it from here
				suppressClick = true;
				if (!commit || (times.startMs === ev.startMs && times.endMs === ev.endMs)) {
					this.render(); // snap back
					return;
				}
				void this.plugin.moveEvent(ev, times.startMs, times.endMs);
			};
			const up = () => finish(true);
			const onKey = (ke: KeyboardEvent) => {
				if (ke.key === "Escape") {
					ke.preventDefault();
					ke.stopPropagation();
					finish(false);
				}
			};
			document.addEventListener("pointermove", move);
			document.addEventListener("pointerup", up);
			document.addEventListener("keydown", onKey, true);
		});
	}

	/** Drag across empty grid to draft a timed event there; double-click for a
	 *  half-hour slot. Presses on blocks never reach this (target check). */
	private attachSlotGesture(col: HTMLElement, key: string) {
		col.addEventListener("dblclick", (e) => {
			if (e.target !== col || !this.plugin.anyWritable()) return;
			const rect = col.getBoundingClientRect();
			const min = Math.max(0, Math.min(1410, snapMin(((e.clientY - rect.top) / HOUR_H) * 60, 30)));
			this.openEventModal(null, msOfKey(key) + min * 60000, msOfKey(key) + (min + 30) * 60000, false);
		});
		col.addEventListener("pointerdown", (pd: PointerEvent) => {
			if (pd.button !== 0 || pd.pointerType !== "mouse" || pd.target !== col || !this.plugin.anyWritable()) return;
			const rect = col.getBoundingClientRect();
			const anchorMin = Math.max(0, Math.min(1440, snapMin(((pd.clientY - rect.top) / HOUR_H) * 60, 15)));
			let ghost: HTMLElement | null = null;
			let a = anchorMin;
			let b = anchorMin;
			const move = (mv: PointerEvent) => {
				const cur = Math.max(0, Math.min(1440, snapMin(((mv.clientY - rect.top) / HOUR_H) * 60, 15)));
				if (!ghost) {
					if (Math.abs(mv.clientY - pd.clientY) < 5) return;
					ghost = col.createDiv("pcal-ghost");
				}
				a = Math.min(anchorMin, cur);
				b = Math.max(a + 15, Math.max(anchorMin, cur));
				ghost.style.top = `${(a / 60) * HOUR_H}px`;
				ghost.style.height = `${((b - a) / 60) * HOUR_H}px`;
				ghost.setText(`${fmtClock(a, this.plugin.settings.use24h, true)} - ${fmtClock(Math.min(b, 1439), this.plugin.settings.use24h, true)}`);
			};
			const up = () => {
				document.removeEventListener("pointermove", move);
				document.removeEventListener("pointerup", up);
				if (!ghost) return; // a plain click on empty grid stays inert
				ghost.remove();
				this.openEventModal(null, msOfKey(key) + a * 60000, msOfKey(key) + b * 60000, false);
			};
			document.addEventListener("pointermove", move);
			document.addEventListener("pointerup", up);
		});
	}

	/* ---------------- event card ---------------- */

	private closeCard() {
		this.cardCleanup?.();
		this.cardCleanup = null;
		this.cardEl?.remove();
		this.cardEl = null;
	}

	private openCard(ev: PCEvent, anchor: HTMLElement) {
		if (!anchor.isConnected) return; // a drop just re-rendered under this click
		this.closeCard();
		const s = this.plugin.settings;
		const card = document.body.createDiv("pcal-card");
		this.cardEl = card;
		card.style.setProperty("--pcal-ev-color", ev.color ?? "var(--interactive-accent)");

		// context first, like Outlook's peek: which calendar this event lives on
		if (ev.calendarName) {
			const ctx = card.createDiv("pcal-card-context");
			ctx.createSpan("pcal-card-cal-dot");
			ctx.createSpan({ text: ev.calendarName + (ev.tentative ? " · tentative" : "") + (ev.declined ? " · declined" : "") });
		}
		const head = card.createDiv("pcal-card-head");
		head.createDiv("pcal-card-bar");
		const ht = head.createDiv("pcal-card-headtext");
		ht.createDiv({ cls: "pcal-card-title", text: ev.title });

		const people = dedupePeople(ev.organizer, ev.attendees);

		// the two actions that matter ride right under the title; everything
		// occasional lives behind the ellipsis so the card stays calm
		const btns = card.createDiv("pcal-card-btns");
		if (ev.joinUrl) {
			const join = btns.createEl("button", { cls: "mod-cta", text: "Join" });
			join.addEventListener("click", () => window.open(ev.joinUrl, "_blank"));
		}
		const note = btns.createEl("button", { text: this.plugin.noteExistsFor(ev) ? "Open note" : "Create note" });
		note.addEventListener("click", () => {
			this.closeCard();
			void this.plugin.openEventNote(ev);
		});
		if (ev.canEdit && !ev.notePath) {
			const edit = btns.createEl("button", { text: "Edit" });
			edit.addEventListener("click", () => {
				this.closeCard();
				if (ev.recurring && ev.seriesId) {
					new SeriesChoiceModal(this.app, `Edit "${ev.title}"`, (scope) => {
						if (scope === "occurrence") this.openEventModal(ev, ev.startMs, ev.endMs, ev.allDay);
						else
							void this.plugin.loadSeriesMaster(ev).then((master) => {
								if (master) this.openEventModal(master, master.startMs, master.endMs, master.allDay);
							});
					}).open();
				} else this.openEventModal(ev, ev.startMs, ev.endMs, ev.allDay);
			});
		}
		const moreItems: { title: string; warning?: boolean; onClick: () => void }[] = [];
		if (!ev.notePath && this.plugin.assistantNewMeeting())
			moreItems.push({
				title: "Capture meeting",
				onClick: () => {
					this.closeCard();
					this.plugin.captureMeeting(ev);
				},
			});
		if (!ev.notePath && (ev.attendeeDetail ?? []).some((a) => a.email))
			moreItems.push({
				title: "Email attendees",
				onClick: () => {
					this.closeCard();
					const to = (ev.attendeeDetail ?? [])
						.map((a) => a.email)
						.filter((e): e is string => !!e)
						.join(", ");
					const preferAccountId = ev.sourceId.startsWith("m365:") ? ev.sourceId.split(":")[1] : undefined;
					new RichComposeModal(this.app, this.plugin, { mode: "new", to, subject: ev.title, preferAccountId }).open();
				},
			});
		if (!ev.notePath && this.plugin.anyWritable())
			moreItems.push({
				title: "Duplicate",
				onClick: () => {
					this.closeCard();
					const invites = (ev.attendeeDetail ?? [])
						.map((a) => (a.name && a.email ? `${a.name} <${a.email}>` : a.email ?? ""))
						.filter(Boolean)
						.join(", ");
					this.openEventModal(null, ev.startMs, ev.endMs, ev.allDay, { title: ev.title, location: ev.location, invites });
				},
			});
		if (ev.url)
			moreItems.push({
				title: "Open original",
				onClick: () => window.open(ev.url, "_blank"),
			});
		if (ev.canEdit && !ev.notePath)
			moreItems.push({
				title: "Delete",
				warning: true,
				onClick: () => {
					this.closeCard();
					const notified = people.length ? " Attendees are notified of the cancellation." : "";
					const confirmOccurrence = () =>
						new ConfirmModal(
							this.app,
							ev.recurring ? `Cancel this occurrence of "${ev.title}"?` : `Delete "${ev.title}"?`,
							(ev.recurring ? "Only this occurrence is removed." : "The event is removed from your calendar.") + notified + " It lands in your mailbox's Deleted Items, so it is recoverable there.",
							ev.recurring ? "Cancel occurrence" : "Delete",
							() => void this.plugin.deleteCalEvent(ev)
						).open();
					if (ev.recurring && ev.seriesId) {
						new SeriesChoiceModal(this.app, `Delete "${ev.title}"`, (scope) => {
							if (scope === "occurrence") confirmOccurrence();
							else
								new ConfirmModal(
									this.app,
									`Delete the whole series "${ev.title}"?`,
									"Every occurrence is removed." + notified + " It lands in your mailbox's Deleted Items, so it is recoverable there.",
									"Delete series",
									() => void this.plugin.deleteCalEvent({ ...ev, id: ev.seriesId as string, recurring: false })
								).open();
						}).open();
					} else confirmOccurrence();
				},
			});
		if (moreItems.length) {
			const moreBtn = btns.createEl("button", { cls: "pcal-card-more", attr: { "aria-label": "More actions" } });
			setIcon(moreBtn, "more-horizontal");
			moreBtn.addEventListener("click", (e) => {
				const menu = new Menu();
				for (const it of moreItems)
					menu.addItem((i) => {
						i.setTitle(it.title).onClick(it.onClick);
						if (it.warning) i.setWarning(true);
					});
				menu.showAtMouseEvent(e);
			});
		}

		const meta = card.createDiv("pcal-card-meta");
		const metaRow = (icon: string, text: string): HTMLElement => {
			const row = meta.createDiv("pcal-card-row");
			const ic = row.createSpan("pcal-card-icon");
			setIcon(ic, icon);
			row.createSpan({ text });
			return row;
		};
		const span = eventDaySpan(ev);
		// a multi-day range already names its days; prefixing the heading again reads twice
		const whenRow = metaRow("clock", span.startKey === span.endKey ? `${fmtDayHeading(span.startKey)} · ${fmtEventRange(ev, s.use24h)}` : fmtEventRange(ev, s.use24h));
		if (ev.recurring) {
			const rep = whenRow.createSpan("pcal-card-repeat");
			setIcon(rep, "repeat");
		}
		// a location that only names the meeting platform is the join button again
		const loc = ev.location?.trim() ?? "";
		if (loc && !(ev.joinUrl && /^(microsoft teams|zoom|google meet|webex)/i.test(loc))) metaRow("map-pin", loc);
		// people are links into their pages, so a meeting's attendees connect
		// to the vault's person hubs in one click
		const personRow = (icon: string, names: string[], extra?: string) => {
			const row = meta.createDiv("pcal-card-row");
			const ic = row.createSpan("pcal-card-icon");
			setIcon(ic, icon);
			const wrap = row.createSpan("pcal-card-people");
			names.forEach((n, idx) => {
				if (idx) wrap.appendText(", ");
				const link = wrap.createSpan({ cls: "pcal-person-link", text: n });
				link.addEventListener("click", () => {
					this.closeCard();
					void this.plugin.openPersonPage(n);
				});
			});
			if (extra) wrap.appendText(extra);
		};
		if (ev.organizer) personRow("user", [ev.organizer]);
		if (people.length) personRow("users", people.slice(0, 6), people.length > 6 ? ` +${people.length - 6}` : undefined);

		if (ev.canRsvp) {
			const rsvp = card.createDiv("pcal-card-rsvp");
			rsvp.createSpan({ cls: "pcal-rsvp-label", text: "Going?" });
			const opts: { r: "accepted" | "tentative" | "declined"; label: string }[] = [
				{ r: "accepted", label: "Yes" },
				{ r: "tentative", label: "Maybe" },
				{ r: "declined", label: "No" },
			];
			for (const o of opts) {
				const b = rsvp.createEl("button", { cls: "pcal-rsvp-btn", text: o.label });
				b.toggleClass("is-active", ev.myResponse === o.r);
				b.addEventListener("click", () => {
					this.closeCard();
					void this.plugin.rsvpEvent(ev, o.r);
				});
			}
		}

		const d = stripMeetingBoilerplate(ev.description ?? "");
		if (d) card.createDiv({ cls: "pcal-card-desc", text: d.length > 280 ? d.slice(0, 280).trimEnd() + "..." : d });

		// place near the anchor, clamped to the viewport, flipping above when
		// there is no room below (same mechanics as the settings help popover)
		const r = anchor.getBoundingClientRect();
		card.style.left = Math.max(8, Math.min(r.left, window.innerWidth - card.offsetWidth - 8)) + "px";
		const below = r.bottom + 6;
		card.style.top = (below + card.offsetHeight > window.innerHeight - 8 ? Math.max(8, r.top - card.offsetHeight - 6) : below) + "px";

		const onDocDown = (e: MouseEvent) => {
			if (e.target instanceof Node && (card.contains(e.target) || anchor.contains(e.target))) return;
			this.closeCard();
		};
		const onKey = (e: KeyboardEvent) => {
			if (e.key === "Escape") this.closeCard();
		};
		const onScroll = (e: Event) => {
			if (e.target instanceof Node && card.contains(e.target)) return;
			this.closeCard();
		};
		document.addEventListener("pointerdown", onDocDown, true);
		document.addEventListener("keydown", onKey, true);
		document.addEventListener("scroll", onScroll, true);
		this.cardCleanup = () => {
			document.removeEventListener("pointerdown", onDocDown, true);
			document.removeEventListener("keydown", onKey, true);
			document.removeEventListener("scroll", onScroll, true);
		};
	}
}

/* ---------------- modals ---------------- */

/** "This occurrence or the whole series?" Occurrence leads, like Outlook. */
class SeriesChoiceModal extends Modal {
	constructor(
		app: App,
		private heading: string,
		private onPick: (scope: "occurrence" | "series") => void
	) {
		super(app);
	}
	onOpen() {
		this.titleEl.setText(this.heading);
		this.contentEl.createEl("p", { cls: "pcal-modal-desc", text: "This event repeats." });
		const btns = this.contentEl.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "The whole series" }).addEventListener("click", () => {
			this.close();
			this.onPick("series");
		});
		const occ = btns.createEl("button", { text: "This occurrence", cls: "mod-cta" });
		occ.addEventListener("click", () => {
			this.close();
			this.onPick("occurrence");
		});
		occ.focus();
	}
	onClose() {
		this.contentEl.empty();
	}
}

class ConfirmModal extends Modal {
	constructor(
		app: App,
		private heading: string,
		private body: string,
		private confirmText: string,
		private onConfirm: () => void
	) {
		super(app);
	}
	onOpen() {
		this.titleEl.setText(this.heading);
		this.contentEl.createEl("p", { cls: "pcal-modal-desc", text: this.body });
		const btns = this.contentEl.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		const go = btns.createEl("button", { text: this.confirmText, cls: "mod-warning" });
		go.addEventListener("click", () => {
			this.close();
			this.onConfirm();
		});
		go.focus();
	}
	onClose() {
		this.contentEl.empty();
	}
}

/** Create or edit a calendar event. Editing never touches the description:
 *  the fetched bodyPreview is a truncated copy, and writing it back would
 *  overwrite the real body with its own stub. */
class EventModal extends Modal {
	private etitle: string;
	private allDay: boolean;
	private draftStart: number;
	private draftEnd: number;
	private location: string;
	private description = "";
	private invites: string;
	private invitesInitial: string;
	private repeat: RepeatKind = "none";
	private showAs: "busy" | "free" | "tentative";
	private targetKey: string | null;
	private fieldsEl!: HTMLElement;
	private fieldsHooked = false;
	private availEl: HTMLElement | null = null;
	private availData: { windowStart: number; windowEnd: number; rows: { email: string; runs: BusyRun[]; error: string | null }[] } | null = null;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private ev: PCEvent | null,
		startMs: number,
		endMs: number,
		allDay: boolean,
		prefill?: { title?: string; location?: string; invites?: string }
	) {
		super(app);
		this.etitle = ev ? (ev.title === "(no title)" ? "" : ev.title) : prefill?.title ?? "";
		this.allDay = allDay;
		this.draftStart = startMs;
		this.draftEnd = Math.max(endMs, startMs + (allDay ? 86400000 : 60000 * 15));
		this.location = ev?.location ?? prefill?.location ?? "";
		this.invites = ev
			? (ev.attendeeDetail ?? [])
					.map((a) => (a.name && a.email ? `${a.name} <${a.email}>` : a.email ?? ""))
					.filter(Boolean)
					.join(", ")
			: prefill?.invites ?? "";
		this.invitesInitial = this.invites;
		this.showAs = ev ? (ev.transparent ? "free" : ev.tentative ? "tentative" : "busy") : "busy";
		this.targetKey = this.plugin.defaultWriteTarget()?.key ?? null;
	}

	onOpen() {
		this.titleEl.setText(this.ev ? "Edit event" : "New event");
		const c = this.contentEl;
		c.addClass("pcal-event-modal");
		let titleInput: HTMLInputElement | null = null;
		new Setting(c).setName("Title").addText((t) => {
			titleInput = t.inputEl;
			t.setPlaceholder("Event title").setValue(this.etitle).onChange((v) => (this.etitle = v));
			t.inputEl.addEventListener("keydown", (e) => {
				if (e.key === "Enter") {
					e.preventDefault();
					void this.save();
				}
			});
		});
		const targets = this.plugin.writableTargets();
		if (!this.ev && targets.length > 1) {
			new Setting(c).setName("Calendar").addDropdown((d) => {
				for (const t of targets) d.addOption(t.key, t.label);
				d.setValue(this.targetKey ?? targets[0].key).onChange((v) => (this.targetKey = v));
			});
		}
		new Setting(c).setName("All day").addToggle((t) =>
			t.setValue(this.allDay).onChange((v) => {
				if (v) {
					const lastKey = keyOfMs(Math.max(this.draftEnd - 1, this.draftStart));
					this.draftStart = msOfKey(keyOfMs(this.draftStart));
					this.draftEnd = msOfKey(addDays(lastKey, 1));
				} else {
					const day = keyOfMs(this.draftStart);
					this.draftStart = msOfKey(day) + 9 * 3600000;
					this.draftEnd = this.draftStart + 30 * 60000;
				}
				this.allDay = v;
				this.renderTimeFields();
			})
		);
		this.fieldsEl = c.createDiv();
		this.renderTimeFields();
		new Setting(c).setName("Location").addText((t) => t.setValue(this.location).onChange((v) => (this.location = v)));
		new Setting(c)
			.setName("Invite")
			.setDesc("Email addresses, comma separated. Invitations go out when you save.")
			.addText((t) => t.setPlaceholder("ana@x.com, Bob <bob@x.com>").setValue(this.invites).onChange((v) => (this.invites = v)));
		new Setting(c).setName("Show as").addDropdown((d) =>
			d
				.addOptions({ busy: "Busy", free: "Free", tentative: "Tentative" })
				.setValue(this.showAs)
				.onChange((v) => (this.showAs = v as "busy" | "free" | "tentative"))
		);
		new Setting(c)
			.setName("Availability")
			.setDesc("Invitees' free/busy around this time (same organization, Microsoft 365).")
			.addButton((b) => b.setButtonText("Check").onClick(() => void this.checkAvailability()));
		this.availEl = c.createDiv("pcal-avail");
		if (!this.ev) {
			new Setting(c).setName("Repeat").addDropdown((d) =>
				d
					.addOptions({ none: "Does not repeat", daily: "Daily", weekdays: "Every weekday", weekly: "Weekly", monthly: "Monthly", yearly: "Yearly" })
					.setValue(this.repeat)
					.onChange((v) => (this.repeat = v as RepeatKind))
			);
			new Setting(c).setName("Description").addTextArea((t) => {
				t.setValue(this.description).onChange((v) => (this.description = v));
				t.inputEl.rows = 3;
			});
		}
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: this.ev ? "Save" : "Create", cls: "mod-cta" }).addEventListener("click", () => void this.save());
		window.setTimeout(() => titleInput?.focus(), 20);
	}

	/** The date/time rows, rebuilt when the all-day toggle flips. Start edits
	 *  carry the duration along; end edits set the end directly. */
	private renderTimeFields() {
		const host = this.fieldsEl;
		host.empty();
		const pad = (n: number) => String(n).padStart(2, "0");
		const timeVal = (min: number) => `${pad(Math.floor(min / 60))}:${pad(min % 60)}`;
		// change handlers only update state: Chromium fires change per typed
		// digit on date/time inputs, and rebuilding here would destroy the
		// input between two digits and eat the second one. The rebuild that
		// syncs dependent rows waits until focus leaves the whole group.
		if (!this.fieldsHooked) {
			this.fieldsHooked = true;
			host.addEventListener("focusout", (e) => {
				const to = e.relatedTarget;
				if (to instanceof Node && host.contains(to)) return; // tabbing between the fields
				this.renderTimeFields();
			});
		}
		const dateInput = (row: Setting, value: string, onPick: (key: string) => void) => {
			const el = row.controlEl.createEl("input", { attr: { type: "date" } });
			el.value = value;
			el.addEventListener("change", () => {
				if (/^\d{4}-\d{2}-\d{2}$/.test(el.value)) onPick(el.value);
				this.renderAvail();
			});
		};
		const timeInput = (row: Setting, value: number, onPick: (min: number) => void) => {
			const el = row.controlEl.createEl("input", { attr: { type: "time" } });
			el.value = timeVal(value);
			el.addEventListener("change", () => {
				const m = el.value.match(/^(\d{2}):(\d{2})$/);
				if (m) onPick(+m[1] * 60 + +m[2]);
				this.renderAvail();
			});
		};
		const dur = this.draftEnd - this.draftStart;
		const start = new Setting(host).setName(this.allDay ? "First day" : "Starts");
		dateInput(start, keyOfMs(this.draftStart), (key) => {
			this.draftStart = msOfKey(key) + (this.allDay ? 0 : minutesOfMs(this.draftStart) * 60000);
			this.draftEnd = this.draftStart + dur;
		});
		if (!this.allDay)
			timeInput(start, minutesOfMs(this.draftStart), (min) => {
				this.draftStart = msOfKey(keyOfMs(this.draftStart)) + min * 60000;
				this.draftEnd = this.draftStart + dur;
			});
		const end = new Setting(host).setName(this.allDay ? "Last day" : "Ends");
		if (this.allDay) {
			dateInput(end, keyOfMs(this.draftEnd - 1), (key) => {
				this.draftEnd = Math.max(msOfKey(addDays(key, 1)), this.draftStart + 86400000);
			});
		} else {
			// a midnight end belongs to the next day, which the date field makes visible
			if (keyOfMs(this.draftEnd) !== keyOfMs(this.draftStart))
				dateInput(end, keyOfMs(this.draftEnd), (key) => {
					this.draftEnd = msOfKey(key) + minutesOfMs(this.draftEnd) * 60000;
				});
			timeInput(end, minutesOfMs(this.draftEnd), (min) => {
				this.draftEnd = msOfKey(keyOfMs(this.draftEnd)) + min * 60000;
			});
		}
		this.renderAvail(); // the draft marker and verdicts follow the times
	}

	/** Ask Graph for the invitees' free/busy over the draft's day and draw it. */
	private async checkAvailability() {
		const targets = this.plugin.writableTargets();
		const t = (this.targetKey ? targets.find((x) => x.key === this.targetKey) : null) ?? this.plugin.defaultWriteTarget();
		const def = t?.def;
		if (!def || def.kind !== "m365") {
			new Notice("Power Desk: availability needs a Microsoft 365 calendar.");
			return;
		}
		const emails = parseAttendeeInput(this.invites).map((a) => a.email);
		if (!emails.length) {
			new Notice("Power Desk: add invitees first.");
			return;
		}
		const s = this.plugin.settings;
		const day = keyOfMs(this.draftStart);
		const fromH = Math.min(s.freeFromHour, Math.floor(minutesOfMs(this.draftStart) / 60));
		const toH = Math.min(24, Math.max(s.freeToHour, Math.ceil((minutesOfMs(Math.max(this.draftEnd - 1, this.draftStart)) + 1) / 60)));
		const windowStart = msOfKey(day) + fromH * 3600000;
		const windowEnd = msOfKey(day) + toH * 3600000;
		this.availEl?.setText("Checking...");
		const rows = await this.plugin.checkAvailability(def, emails, windowStart, windowEnd);
		if (!rows) {
			this.availEl?.setText("");
			return;
		}
		this.availData = { windowStart, windowEnd, rows };
		this.renderAvail();
	}

	/** Bars per invitee with the draft slot overlaid; re-run whenever the
	 *  times change so the verdicts follow the draft. */
	private renderAvail() {
		const host = this.availEl;
		const d = this.availData;
		if (!host) return;
		if (!d) return;
		host.empty();
		if (keyOfMs(this.draftStart) !== keyOfMs(d.windowStart)) {
			host.createDiv({ cls: "pcal-avail-summary", text: "The date changed. Check again." });
			return;
		}
		const span = d.windowEnd - d.windowStart;
		const pct = (ms: number) => `${Math.max(0, Math.min(100, ((ms - d.windowStart) / span) * 100))}%`;
		const use24h = this.plugin.settings.use24h;
		let anyConflict = false;
		for (const row of d.rows) {
			const r = host.createDiv("pcal-avail-row");
			r.createDiv({ cls: "pcal-avail-name", text: row.email.split("@")[0] });
			const track = r.createDiv("pcal-avail-track");
			if (row.error) {
				track.addClass("is-unknown");
				track.setText("no visibility");
				r.createDiv({ cls: "pcal-avail-verdict", text: "?" });
				continue;
			}
			for (const run of row.runs) {
				const b = track.createDiv(`pcal-avail-block is-${run.kind}`);
				b.style.left = pct(run.startMs);
				b.style.width = `calc(${pct(run.endMs)} - ${pct(run.startMs)})`;
			}
			const mark = track.createDiv("pcal-avail-slot");
			mark.style.left = pct(this.draftStart);
			mark.style.width = `calc(${pct(Math.min(this.draftEnd, d.windowEnd))} - ${pct(this.draftStart)})`;
			const worst = slotConflict(row.runs, this.draftStart, this.draftEnd);
			if (worst && worst !== "tentative") anyConflict = true;
			r.createDiv({
				cls: "pcal-avail-verdict" + (worst ? (worst === "tentative" ? " is-soft" : " is-bad") : " is-ok"),
				text: worst ? (worst === "tentative" ? "tentative" : "busy") : "free",
			});
		}
		host.createDiv({
			cls: "pcal-avail-summary",
			text: `${fmtTimeOfMs(d.windowStart, use24h, true)} - ${fmtTimeOfMs(d.windowEnd, use24h, true)} · ` + (anyConflict ? "Someone is busy at this time." : "Everyone is free at this time."),
		});
	}

	private async save() {
		if (this.draftEnd <= this.draftStart) {
			new Notice("Power Desk: the end must come after the start.");
			return;
		}
		// attendees ride along only when they changed, so an untouched invite
		// list can never resend or clobber a meeting's roster
		const parsed = parseAttendeeInput(this.invites);
		const attendees = this.ev ? (this.invites.trim() !== this.invitesInitial.trim() ? parsed : null) : parsed.length ? parsed : null;
		const draft: EventDraft = {
			title: this.etitle,
			startMs: this.draftStart,
			endMs: this.draftEnd,
			allDay: this.allDay,
			location: this.location,
			description: this.ev ? undefined : this.description || undefined,
			attendees,
			repeat: this.ev ? undefined : this.repeat,
			showAs: this.showAs,
		};
		const ok = this.ev ? await this.plugin.updateCalEvent(this.ev, draft) : await this.plugin.createEventAt(this.targetKey, draft);
		if (ok) this.close();
	}

	onClose() {
		this.contentEl.empty();
	}
}

class DeviceCodeModal extends Modal {
	waiting = true;
	constructor(
		app: App,
		private dc: DeviceCode
	) {
		super(app);
	}
	onOpen() {
		this.titleEl.setText("Connect Microsoft 365");
		const c = this.contentEl;
		c.createEl("p", {
			cls: "pcal-modal-desc",
			text: "Sign in with your Microsoft account so Power Desk can read your calendar. Open the page, enter the code, and approve. This window finishes automatically.",
		});
		c.createEl("div", { cls: "pcal-devicecode", text: this.dc.user_code });
		const row = c.createDiv({ cls: "pcal-modal-btns pcal-left" });
		row.createEl("button", { text: "Copy code" }).addEventListener("click", () => {
			void navigator.clipboard.writeText(this.dc.user_code);
			new Notice("Power Desk: code copied.");
		});
		row.createEl("button", { text: "Open sign-in page", cls: "mod-cta" }).addEventListener("click", () => window.open(this.dc.verification_uri, "_blank"));
		c.createEl("p", { cls: "pcal-modal-desc pcal-devicecode-url", text: this.dc.verification_uri });
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
	}
	onClose() {
		this.waiting = false;
		this.contentEl.empty();
	}
}

/** Compose a mail: sent from one of the connected mailboxes when possible
 *  (Power Assistant's transport as a fallback), or handed to the system mail
 *  app, which is all Notion Calendar does either. */
const SEARCH_FOLDER_TYPES: { v: string; label: string; param?: string }[] = [
	{ v: "unread", label: "Unread mail" },
	{ v: "flagged", label: "Flagged for follow-up" },
	{ v: "unreadOrFlagged", label: "Unread or flagged" },
	{ v: "important", label: "Important mail" },
	{ v: "toMe", label: "Mail sent directly to me" },
	{ v: "fromPeople", label: "Mail from specific people", param: "People (addresses, comma separated)" },
	{ v: "fromToPeople", label: "Mail from and to specific people", param: "People (addresses, comma separated)" },
	{ v: "toGroup", label: "Mail sent to public groups", param: "Group addresses (comma separated)" },
	{ v: "category", label: "Categorized mail", param: "Category name" },
	{ v: "attachments", label: "Mail with attachments" },
	{ v: "custom", label: "Custom search", param: "Query (like from:ana@x.com subject:invoice)" },
];

/** Outlook's New-search-folder dialog: pick a type from the catalog, fill in
 *  its people or category when it asks, and the folder saves per account. */
class SearchFolderModal extends Modal {
	private name: string;
	private type: string;
	private param: string;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private account: GraphAccount,
		private existing: { id: string; name: string; query: string; type?: string; param?: string } | null,
		private onSave: () => void
	) {
		super(app);
		this.name = existing?.name ?? "";
		this.type = existing?.type ?? (existing ? "custom" : "unread");
		this.param = existing?.param ?? (existing && !existing.type ? existing.query : "");
	}

	onOpen() {
		this.titleEl.setText(this.existing ? "Edit search folder" : "New search folder");
		const c = this.contentEl;
		c.addClass("pcal-searchfolder-modal");
		let nameInput: HTMLInputElement | null = null;
		new Setting(c).setName("Name").addText((t) => {
			nameInput = t.inputEl;
			t.setValue(this.name).onChange((v) => (this.name = v));
		});
		let paramHost: HTMLDivElement;
		const renderParam = () => {
			paramHost.empty();
			const def = SEARCH_FOLDER_TYPES.find((x) => x.v === this.type);
			if (!def?.param) return;
			// "People (addresses, comma separated)" splits into label + hint line
			const m = def.param.match(/^(.*?)\s*\((.*)\)$/);
			const st = new Setting(paramHost).setName(m ? m[1] : def.param);
			if (m) st.setDesc(m[2].charAt(0).toUpperCase() + m[2].slice(1));
			st.addText((t) => t.setValue(this.param).onChange((v) => (this.param = v)));
		};
		new Setting(c).setName("Type").addDropdown((d) => {
			for (const t of SEARCH_FOLDER_TYPES) d.addOption(t.v, t.label);
			d.setValue(this.type).onChange((v) => {
				const prev = SEARCH_FOLDER_TYPES.find((x) => x.v === this.type);
				this.type = v;
				const def = SEARCH_FOLDER_TYPES.find((x) => x.v === v);
				// the name follows the type until the user writes their own
				if (nameInput && def && (!this.name.trim() || this.name === prev?.label)) {
					this.name = def.label;
					nameInput.value = def.label;
				}
				renderParam();
			});
		});
		paramHost = c.createDiv();
		renderParam();
		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			const def = SEARCH_FOLDER_TYPES.find((x) => x.v === this.type);
			if (def?.param && !this.param.trim()) {
				new Notice(`Power Desk: fill in ${def.param.toLowerCase()}.`);
				return;
			}
			const query = searchFolderQuery(this.type, this.param, this.account.label);
			if (!query.trim()) {
				new Notice("Power Desk: this search folder would match nothing.");
				return;
			}
			const name = this.name.trim() || def?.label || "Search";
			const s = this.plugin.settings;
			if (this.existing) {
				s.mailSearchFolders = s.mailSearchFolders.map((x) => (x.id === this.existing?.id ? { ...x, name, query, type: this.type, param: this.param } : x));
			} else {
				s.mailSearchFolders = [...s.mailSearchFolders, { accountId: this.account.id, id: freshId(), name, query, type: this.type, param: this.param }];
			}
			this.plugin.queueSave();
			this.onSave();
			this.close();
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** The full compose window, Outlook's shape: a formatting toolbar, editable
 *  To, Cc, and Subject, the signature visible in the body, and for replies
 *  and forwards a Graph draft carrying the quoted original right in the
 *  editor. Sending patches and sends the draft; discarding deletes it. */
class RichComposeModal extends Modal {
	private draft: DraftMessage | null = null;
	private sent = false;
	private toInput!: HTMLInputElement;
	private ccInput!: HTMLInputElement;
	private subjInput!: HTMLInputElement;
	private editorEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private opts: { mode: "new" | "reply" | "replyAll" | "forward"; mail?: PCMail; to?: string; subject?: string; preferAccountId?: string }
	) {
		super(app);
	}

	onOpen() {
		const { mode, mail } = this.opts;
		this.modalEl.addClass("pcal-compose-window");
		this.titleEl.setText(mode === "new" ? "New mail" : mode === "reply" ? "Reply" : mode === "replyAll" ? "Reply all" : "Forward");
		const c = this.contentEl;
		c.addClass("pcal-compose");
		const fromLabel = mode === "new" ? this.plugin.mailSender(this.opts.preferAccountId)?.label : mail?.accountLabel;
		if (fromLabel) c.createDiv({ cls: "pcal-modal-desc pcal-compose-from", text: `From ${fromLabel}` });

		const row = (label: string, value: string, placeholder?: string): HTMLInputElement => {
			const r = c.createDiv("pcal-compose-row");
			r.createSpan({ cls: "pcal-compose-label", text: label });
			const inp = r.createEl("input", { attr: { type: "text", spellcheck: "false" } });
			inp.value = value;
			if (placeholder) inp.placeholder = placeholder;
			return inp;
		};
		this.toInput = row("To", this.opts.to ?? "", "ana@x.com, bob@y.com");
		this.ccInput = row("Cc", "");
		this.subjInput = row("Subject", this.opts.subject ?? "");

		// the formatting bar keeps the editor's selection because its buttons
		// swallow mousedown before focus can move
		const bar = c.createDiv("pcal-compose-bar");
		const tb = (icon: string, label: string, run: () => void) => {
			const b = bar.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
			setIcon(b, icon);
			b.addEventListener("mousedown", (e) => e.preventDefault());
			b.addEventListener("click", () => {
				run();
				this.editorEl.focus();
			});
		};
		const cmd = (name: string) => () => document.execCommand(name);
		tb("bold", "Bold", cmd("bold"));
		tb("italic", "Italic", cmd("italic"));
		tb("underline", "Underline", cmd("underline"));
		tb("strikethrough", "Strikethrough", cmd("strikeThrough"));
		tb("list", "Bulleted list", cmd("insertUnorderedList"));
		tb("list-ordered", "Numbered list", cmd("insertOrderedList"));
		tb("link", "Link", () => {
			const sel = window.getSelection();
			const range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
			new PromptModal(this.app, "Link", [{ label: "Address", value: "", placeholder: "https://..." }], ([url]) => {
				if (!url.trim()) return;
				this.editorEl.focus();
				if (range && sel) {
					sel.removeAllRanges();
					sel.addRange(range);
				}
				document.execCommand("createLink", false, url.trim());
			}).open();
		});
		tb("eraser", "Clear formatting", cmd("removeFormat"));

		this.editorEl = c.createDiv({ cls: "pcal-compose-editor", attr: { contenteditable: "true" } });
		if (mode === "new") {
			this.editorEl.createEl("p").createEl("br");
			this.editorEl.appendChild(sanitizeHTMLToDom(this.plugin.signatureHtml()));
		} else this.editorEl.createEl("p", { cls: "pcal-compose-loading", text: "Opening the draft..." });

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Discard" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Send", cls: "mod-cta" }).addEventListener("click", () => void this.send());

		if (mode !== "new" && mail) void this.loadDraft(mail, mode);
		else window.setTimeout(() => this.editorEl.focus(), 20);
	}

	private async loadDraft(mail: PCMail, kind: "reply" | "replyAll" | "forward") {
		const draft = await this.plugin.createMailDraft(mail, kind);
		if (!this.contentEl.isConnected) {
			// closed before the draft arrived: leave nothing behind in Drafts
			if (draft) void this.plugin.discardMailDraft(mail.accountId, draft.id);
			return;
		}
		if (!draft) {
			this.close();
			return;
		}
		this.draft = draft;
		this.toInput.value = draft.to.join(", ");
		this.ccInput.value = draft.cc.join(", ");
		this.subjInput.value = draft.subject;
		// A reply or forward draft quotes the original message, so this HTML came
		// from whoever sent it. It goes through Obsidian's sanitizer, the same way
		// the reading pane does it: innerHTML would not run <script>, but it does
		// run inline handlers like <img onerror=...>.
		this.editorEl.empty();
		this.editorEl.createEl("p").createEl("br");
		this.editorEl.appendChild(sanitizeHTMLToDom(`${this.plugin.signatureHtml()}${draft.bodyHtml}`));
		window.setTimeout(() => {
			this.editorEl.focus();
			const sel = window.getSelection();
			if (sel && this.editorEl.firstChild) {
				const r = document.createRange();
				r.setStart(this.editorEl.firstChild, 0);
				r.collapse(true);
				sel.removeAllRanges();
				sel.addRange(r);
			}
		}, 20);
	}

	private async send() {
		const split = (s: string) =>
			s
				.split(/[,;]+/)
				.map((x) => x.trim())
				.filter(Boolean);
		const to = split(this.toInput.value);
		const cc = split(this.ccInput.value);
		const subject = this.subjInput.value.trim() || "(no subject)";
		const html = this.editorEl.innerHTML;
		if (!to.length) {
			new Notice("Power Desk: enter at least one recipient.");
			return;
		}
		if (this.opts.mode === "new") {
			const sender = this.plugin.mailSender(this.opts.preferAccountId);
			if (!sender) {
				new Notice("Power Desk: no account can send mail.");
				return;
			}
			try {
				await sender.send({ to, cc, subject, html });
				new Notice("Power Desk: mail sent.");
				this.sent = true;
				this.close();
			} catch (e) {
				new Notice("Power Desk: " + (e instanceof Error ? e.message : String(e)));
			}
			return;
		}
		if (!this.draft || !this.opts.mail) return;
		const ok = await this.plugin.sendMailDraft(this.opts.mail.accountId, this.draft.id, { subject, html, to, cc });
		if (ok) {
			this.sent = true;
			this.close();
		}
	}

	onClose() {
		if (this.draft && !this.sent && this.opts.mail) void this.plugin.discardMailDraft(this.opts.mail.accountId, this.draft.id);
		this.contentEl.empty();
	}
}

/** Adding Microsoft accounts, as a wizard: pick the account kind, settle the
 *  app registration that signs it in (created inline when missing), watch the
 *  device code get approved, and finish on a verified summary that offers the
 *  next account. Failures render inside the step with their fix instead of as
 *  vanishing notices, and a later personal account can reuse the app an
 *  earlier one registered. */
class GraphAccountWizard extends Modal {
	private stepNo: 1 | 2 | 3 | 4 = 1;
	private kind: "work" | "personal" = "work";
	private useOwnApp = false;
	private creating = false;
	private clientId = "";
	private tenant = "";
	private signingIn = false;
	private closed = false;
	private connected: GraphAccount | null = null;
	private codeEl: HTMLElement | null = null;
	private statusEl: HTMLElement | null = null;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Add a Microsoft 365 account");
		this.render();
	}

	onClose() {
		this.closed = true; // ends any device-code poll still waiting
		this.contentEl.empty();
	}

	private render() {
		const c = this.contentEl;
		c.empty();
		this.codeEl = this.statusEl = null;
		const titles = ["Account type", "App registration", "Sign in", "Connected"];
		c.createDiv({ cls: "pcal-wizard-progress", text: `Step ${this.stepNo} of 4 · ${titles[this.stepNo - 1]}` });
		if (this.stepNo === 1) this.renderKind(c);
		else if (this.stepNo === 2) this.renderApp(c);
		else if (this.stepNo === 3) this.renderSignIn(c);
		else this.renderDone(c);
	}

	private footer(c: HTMLElement, buttons: { label: string; cta?: boolean; onClick: () => void }[]) {
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		for (const b of buttons) btns.createEl("button", { text: b.label, cls: b.cta ? "mod-cta" : "" }).addEventListener("click", b.onClick);
	}

	/** The app an earlier personal account signed in with, if any; a second
	 *  personal account rides it silently instead of asking the user to care
	 *  about registrations at all. */
	private reusablePersonalApp(): string | null {
		const own = this.plugin.settings.graphAccounts.filter((x) => x.clientId.trim() && x.clientId.trim() !== this.plugin.effectiveClientId());
		const hit = own.find((x) => ["consumers", "common"].includes(x.tenant.trim().toLowerCase())) ?? own[0];
		return hit ? hit.clientId.trim() : null;
	}

	private card(c: HTMLElement, title: string, desc: string, pick: () => void) {
		const el = c.createDiv({ cls: "pcal-wizard-card" });
		el.createDiv({ cls: "pcal-wizard-card-title", text: title });
		el.createDiv({ cls: "pcal-guide-text", text: desc });
		el.addEventListener("click", pick);
	}

	/* ---- step 1: account type ---- */
	private renderKind(c: HTMLElement) {
		c.createEl("p", { cls: "pcal-modal-desc", text: "Sign-in happens in your browser with a device code; the plugin never sees a password. Which kind of account is this?" });
		this.card(c, "Work or school", "An account from an organization (name@company.com). Signs in through the shared app registration.", () => {
			this.kind = "work";
			this.useOwnApp = false;
			this.creating = false;
			this.clientId = "";
			this.tenant = "";
			this.stepNo = 2;
			this.render();
		});
		this.card(c, "Personal", "outlook.com, hotmail.com, live.com, msn.com. Signs in through its own app registration, set up once.", () => {
			this.kind = "personal";
			this.useOwnApp = true;
			this.creating = false;
			const prior = this.reusablePersonalApp();
			this.clientId = prior ?? "";
			// 'common' issues devicelogin codes with the full account picker;
			// 'consumers' codes only redeem on the /link page, which silently
			// continues as whoever the browser already signed in
			this.tenant = "common";
			this.stepNo = prior ? 3 : 2;
			this.render();
		});
		this.footer(c, [{ label: "Cancel", onClick: () => this.close() }]);
	}

	/* ---- step 2: app registration ---- */
	private renderApp(c: HTMLElement) {
		if (this.kind === "work" && !this.useOwnApp && this.plugin.effectiveClientId()) {
			c.createEl("p", {
				cls: "pcal-modal-desc",
				text: (this.plugin.usingSiblingApp() ? "Sign-in goes through Power Assistant's app registration, borrowed automatically." : "Sign-in goes through the app registration configured in settings.") + " Nothing to configure; continue straight to the sign-in. Every additional work account connects the same way.",
			});
			new Setting(c)
				.setName("Use a different app registration")
				.setDesc("Only for an account whose organization requires its own app.")
				.addToggle((t) =>
					t.setValue(false).onChange(() => {
						this.useOwnApp = true;
						this.render();
					})
				);
			this.footer(c, [
				{ label: "Back", onClick: () => { this.stepNo = 1; this.render(); } },
				{ label: "Continue to sign-in", cta: true, onClick: () => { this.stepNo = 3; this.render(); } },
			]);
			return;
		}
		if (this.kind === "work" && !this.useOwnApp) {
			// no shared app anywhere yet: the one created here lands in settings
			// and every later work account reuses it
			c.createEl("p", { cls: "pcal-modal-desc", text: "One-time setup: Microsoft sign-ins need an app registration, created once in the Azure portal. It is saved in settings and every later work account reuses it." });
			renderAzureAppSteps(c, false, (id) => {
				this.plugin.settings.graphClientId = id;
				this.plugin.queueSave();
				this.plugin.refreshSettingsTab?.();
				this.stepNo = 3;
				this.render();
			}, "Sign-in starts as soon as the ID lands here.");
			this.footer(c, [{ label: "Back", onClick: () => { this.stepNo = 1; this.render(); } }]);
			return;
		}
		// personal, or a work account with its own app: it lives on the account row
		if (this.creating) {
			renderAzureAppSteps(c, this.kind === "personal", (id) => {
				this.clientId = id;
				if (!this.tenant.trim()) this.tenant = "common";
				this.creating = false;
				this.stepNo = 3;
				this.render();
			}, "Sign-in starts as soon as the ID lands here.");
			this.footer(c, [{ label: "Back", onClick: () => { this.creating = false; this.render(); } }]);
			return;
		}
		if (this.kind === "personal") {
			c.createEl("p", { cls: "pcal-modal-desc", text: "A personal account signs in through its own app registration; the shared work app cannot accept it. Create the app now, or paste an ID you already have. One registration signs in any number of personal accounts, so this never comes up again." });
		} else {
			c.createEl("p", { cls: "pcal-modal-desc", text: "This account signs in through its own app registration, remembered on the account." });
		}
		new Setting(c)
			.setName("No app registration yet?")
			.setDesc("About two minutes in the Azure portal; every field is spelled out.")
			.addButton((b) => b.setButtonText("Create it now").setCta().onClick(() => { this.creating = true; this.render(); }));
		new Setting(c)
			.setName("Application (client) ID")
			.setDesc("Already have an app? Its ID is on the Overview page in the Azure portal.")
			.addText((t) => t.setPlaceholder("00000000-0000-...").setValue(this.clientId).onChange((v) => (this.clientId = v)));
		new Setting(c)
			.setName("Tenant")
			.setDesc("'consumers' signs in personal accounts only; 'common' allows both kinds; single-organization apps take their Directory (tenant) ID.")
			.addText((t) => t.setPlaceholder("common").setValue(this.tenant).onChange((v) => (this.tenant = v)));
		this.footer(c, [
			{ label: "Back", onClick: () => { this.stepNo = 1; this.render(); } },
			{
				label: "Continue to sign-in",
				cta: true,
				onClick: () => {
					if (!this.clientId.trim()) {
						new Notice("Power Desk: enter the Application (client) ID, or press Create it now.");
						return;
					}
					this.stepNo = 3;
					this.render();
				},
			},
		]);
	}

	/* ---- step 3: sign in (the device code start doubles as verification:
	 *  a missing app or public-client-off surfaces right here, inline) ---- */
	private renderSignIn(c: HTMLElement) {
		c.createEl("p", { cls: "pcal-modal-desc", text: "Approve the sign-in in your browser: open the page, enter the code, and pick the account. After the code, the page asks which account; choose 'Use another account' to enter an address that is not listed. Enter exactly the code shown here; a code saved by a password manager from an earlier attempt is expired. This screen moves on by itself." });
		this.codeEl = c.createDiv({ cls: "pcal-devicecode", text: "..." });
		const btnRow = c.createDiv({ cls: "pcal-modal-btns pcal-left" });
		this.statusEl = c.createDiv({ cls: "pcal-modal-status", text: "Requesting a code..." });
		this.footer(c, [{ label: "Back", onClick: () => { this.stepNo = 2; this.render(); } }]);
		void this.startSignIn(btnRow);
	}

	private async startSignIn(btnRow: HTMLElement) {
		// a Back-then-Continue can land here while the previous poll drains;
		// wait it out instead of racing it for the connecting flag
		while (this.signingIn) {
			await sleep(400);
			if (this.closed || this.stepNo !== 3) return;
		}
		if (this.plugin.graphConnecting) {
			this.showError("Another sign-in is already in progress; finish or cancel it first.", null);
			return;
		}
		this.signingIn = true;
		this.plugin.graphConnecting = true;
		const own = this.useOwnApp || this.kind === "personal";
		const a: GraphAccount = {
			id: freshId(),
			label: "Microsoft 365",
			clientId: own ? this.clientId.trim() : "",
			tenant: own ? this.tenant.trim() : "",
			refresh: "",
			access: "",
			expiry: 0,
			grantedScope: "",
			calendars: [],
		};
		try {
			const acct = await this.plugin.deviceCodeSession(
				a,
				true,
				(dc) => this.showCode(dc, btnRow),
				() => this.closed || this.stepNo !== 3
			);
			if (acct) {
				this.connected = acct;
				this.stepNo = 4;
				if (!this.closed) this.render();
			} else if (!this.closed && this.stepNo === 3) {
				this.statusEl?.setText("The code expired before the sign-in finished.");
				btnRow.empty();
				btnRow.createEl("button", { text: "Get a new code", cls: "mod-cta" }).addEventListener("click", () => this.render());
			}
		} catch (e) {
			const msg = e instanceof Error ? e.message : String(e);
			this.showError(graphSetupHint(msg), msg);
		} finally {
			this.plugin.graphConnecting = false;
			this.signingIn = false;
		}
	}

	private showCode(dc: DeviceCode, btnRow: HTMLElement) {
		if (this.closed || this.stepNo !== 3) return;
		this.codeEl?.setText(dc.user_code);
		btnRow.empty();
		btnRow.createEl("button", { text: "Copy code" }).addEventListener("click", () => {
			void navigator.clipboard.writeText(dc.user_code);
			new Notice("Power Desk: code copied.");
		});
		btnRow.createEl("button", { text: "Open sign-in page", cls: "mod-cta" }).addEventListener("click", () => window.open(dc.verification_uri, "_blank"));
		this.statusEl?.setText("Waiting for the approval at " + dc.verification_uri + " ...");
	}

	private showError(hint: string | null, raw: string | null) {
		if (this.closed || this.stepNo !== 3 || !this.statusEl) return;
		this.statusEl.empty();
		const box = this.statusEl.createDiv({ cls: "pcal-wizard-error" });
		box.createDiv({ text: hint ?? "The sign-in failed." });
		if (raw) box.createDiv({ cls: "pcal-guide-text", text: raw });
		box.createDiv({ cls: "pcal-guide-text", text: "Press Back to adjust, then try again." });
	}

	/* ---- step 4: connected ---- */
	private renderDone(c: HTMLElement) {
		const a = this.connected;
		c.createEl("p", { cls: "pcal-modal-desc", text: a ? `${a.label} is connected.` : "Connected." });
		if (a) {
			const box = c.createDiv({ cls: "pcal-guide-fields" });
			const line = (label: string, value: string) => {
				const row = box.createDiv({ cls: "pcal-guide-field" });
				row.createDiv({ cls: "pcal-guide-field-name", text: label });
				row.createDiv({ cls: "pcal-guide-field-value", text: value });
			};
			line("Calendars", `${a.calendars.length} found, all enabled; toggles and colors are in settings`);
			line("Mail", this.plugin.canMailAccount(a) ? "inbox available in the Mail view" : "not granted");
			line("App registration", a.clientId.trim() && a.clientId.trim() !== this.plugin.effectiveClientId() ? "this account's own, remembered on it" : "the shared app");
		}
		this.footer(c, [
			{
				label: "Add another account",
				onClick: () => {
					this.stepNo = 1;
					this.kind = "work";
					this.useOwnApp = false;
					this.creating = false;
					this.clientId = "";
					this.tenant = "";
					this.connected = null;
					this.render();
				},
			},
			{ label: "Done", cta: true, onClick: () => this.close() },
		]);
	}
}

const AZURE_APP_CREATE_URL = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/CreateApplicationBlade";

/** Every Azure-portal step for creating an app registration by hand, with the
 *  exact value for each field; rendered into `host` so the settings guide
 *  modal and the Add account wizard share one source of truth. `onId` receives
 *  the pasted Application (client) ID; `tail` closes step 4 with the caller's
 *  own "what happens next" line. */
function renderAzureAppSteps(host: HTMLElement, personal: boolean, onId: (id: string) => void, tail?: string) {
	const step = (n: number, title: string): HTMLElement => {
		const el = host.createDiv({ cls: "pcal-guide-step" });
		el.createDiv({ cls: "pcal-guide-step-num", text: String(n) });
		const body = el.createDiv({ cls: "pcal-guide-step-body" });
		body.createDiv({ cls: "pcal-guide-step-title", text: title });
		return body;
	};
	const text = (h: HTMLElement, t: string) => h.createDiv({ cls: "pcal-guide-text", text: t });

	const s1 = step(1, "Go to the app registration page");
	text(s1, personal ? "Copy this address into your browser and sign in with the personal account whose mail and calendar you want:" : "Copy this address into your browser and sign in with the account that should own the app:");
	const urlRow = s1.createDiv({ cls: "pcal-guide-url-row" });
	urlRow.createDiv({ cls: "pcal-guide-url", text: AZURE_APP_CREATE_URL });
	urlRow.createEl("button", { text: "Copy link" }).addEventListener("click", () => {
		void navigator.clipboard.writeText(AZURE_APP_CREATE_URL);
		new Notice("Power Desk: link copied.");
	});
	text(s1, "Lands on the portal home instead? Search for 'App registrations' at the top and press 'New registration'.");

	const s2 = step(2, "Register the application");
	const fields = s2.createDiv({ cls: "pcal-guide-fields" });
	const field = (name: string, value: string, note?: string, copyable?: boolean) => {
		const row = fields.createDiv({ cls: "pcal-guide-field" });
		row.createDiv({ cls: "pcal-guide-field-name", text: name });
		const v = row.createDiv({ cls: "pcal-guide-field-value" });
		v.createSpan({ cls: "pcal-guide-field-example", text: value });
		if (note) v.createSpan({ cls: "pcal-guide-field-note", text: " (" + note + ")" });
		if (copyable)
			row.createEl("button", { text: "Copy", cls: "pcal-guide-field-copy" }).addEventListener("click", () => {
				void navigator.clipboard.writeText(value);
				new Notice("Power Desk: copied.");
			});
	};
	field("Name", "Obsidian Power Desk", "anything you like", true);
	field("Supported account types", "Any Entra ID Tenant + Personal Microsoft accounts", "older portals word it as 'accounts in any organizational directory and personal Microsoft accounts'");
	field("Redirect URI", "leave empty");
	text(s2, "Press Register.");

	const s3 = step(3, "Allow public client flows");
	text(s3, "In the app's left menu open Manage, then Authentication. Switch 'Allow public client flows' to Yes and press Save. Sign-in fails without this switch.");

	const s4 = step(4, "Paste the ID back");
	text(s4, "The app's Overview page shows Application (client) ID. Paste it here:");
	const idRow = s4.createDiv({ cls: "pcal-guide-url-row" });
	const idInput = idRow.createEl("input", { cls: "pcal-guide-id-input", attr: { placeholder: "00000000-0000-...", spellcheck: "false" } });
	idRow.createEl("button", { text: "Use this ID", cls: "mod-cta" }).addEventListener("click", () => {
		const v = idInput.value.trim();
		if (!v) {
			new Notice("Power Desk: paste the Application (client) ID first.");
			return;
		}
		onId(v);
	});
	if (tail) text(s4, tail);

	if (!personal)
		host.createEl("p", {
			cls: "pcal-modal-desc pcal-guide-note",
			text: "Organizations that block user consent need an admin to grant the delegated Calendars.ReadWrite, Mail.ReadWrite, and Mail.Send permissions under API permissions; everyone else skips that page entirely.",
		});
}

/** The settings-tab walkthrough for the shared app registration. The Add
 *  account wizard embeds the same steps inline for accounts that bring their
 *  own app, so this modal only ever speaks to the shared case. */
class AzureAppGuideModal extends Modal {
	constructor(
		app: App,
		private onId?: (id: string) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Create the app registration");
		const c = this.contentEl;
		c.createEl("p", {
			cls: "pcal-modal-desc",
			text: "Required once, for work and personal accounts alike. No Azure subscription needed, about two minutes: register the app, flip one switch, paste one ID back.",
		});
		renderAzureAppSteps(
			c,
			false,
			(id) => {
				this.onId?.(id);
				this.close();
			},
			"It becomes the app for new sign-ins; press Add account to connect the first one. A personal account gets its own app instead: press Add account and pick Personal, and these same steps appear inside that wizard."
		);
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => this.close());
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Add or edit a CalDAV account: credentials, then discovery fills the
 *  calendar list. Saving without re-discovering keeps known calendars. */
class CaldavAccountModal extends Modal {
	private name: string;
	private serverUrl: string;
	private username: string;
	private password: string;
	private collections: CaldavCollection[];
	private found: DavCollection[] | null = null;
	private busy = false;
	private listEl!: HTMLElement;
	private statusEl!: HTMLElement;

	constructor(
		app: App,
		private existing: CaldavAccount | null,
		private onSave: (account: CaldavAccount) => void
	) {
		super(app);
		this.name = existing?.name ?? "";
		this.serverUrl = existing?.serverUrl ?? "";
		this.username = existing?.username ?? "";
		this.password = existing?.password ?? "";
		this.collections = existing ? structuredClone(existing.collections) : [];
	}

	onOpen() {
		this.titleEl.setText(this.existing ? "Edit CalDAV account" : "Add CalDAV account");
		const c = this.contentEl;
		const desc = c.createEl("p", {
			cls: "pcal-modal-desc",
			text: "iCloud and Fastmail want an app-specific password, not your account password.",
		});
		let serverInput: HTMLInputElement | null = null;
		let nameInput: HTMLInputElement | null = null;
		const PRESETS: Record<string, { name: string; url: string; hint: string }> = {
			icloud: { name: "iCloud", url: "caldav.icloud.com", hint: "Username is your Apple ID; make an app-specific password at appleid.apple.com." },
			fastmail: { name: "Fastmail", url: "caldav.fastmail.com", hint: "Username is your Fastmail address; make an app password in Fastmail's settings." },
			nextcloud: { name: "Nextcloud", url: "https://your.server/remote.php/dav", hint: "Point the URL at your own server's /remote.php/dav." },
		};
		if (!this.existing)
			new Setting(c).setName("Provider").addDropdown((d) => {
				d.addOptions({ custom: "Custom server", icloud: "iCloud", fastmail: "Fastmail", nextcloud: "Nextcloud" });
				d.setValue("custom").onChange((v) => {
					const p = PRESETS[v];
					if (!p) return;
					this.serverUrl = p.url;
					if (serverInput) serverInput.value = p.url;
					if (!this.name.trim() || Object.values(PRESETS).some((x) => x.name === this.name)) {
						this.name = p.name;
						if (nameInput) nameInput.value = p.name;
					}
					desc.setText(p.hint);
				});
			});
		new Setting(c).setName("Name").addText((t) => {
			nameInput = t.inputEl;
			t.setPlaceholder("iCloud").setValue(this.name).onChange((v) => (this.name = v));
		});
		new Setting(c).setName("Server URL").addText((t) => {
			serverInput = t.inputEl;
			t.setPlaceholder("caldav.icloud.com").setValue(this.serverUrl).onChange((v) => (this.serverUrl = v));
		});
		new Setting(c).setName("Username").addText((t) => t.setPlaceholder("you@example.com").setValue(this.username).onChange((v) => (this.username = v)));
		new Setting(c).setName("App password").addText((t) => {
			t.inputEl.type = "password";
			t.setValue(this.password).onChange((v) => (this.password = v));
		});
		const discover = new Setting(c).setName("Calendars").setDesc(this.collections.length ? "" : "Connect to list this account's calendars.");
		discover.addButton((b) =>
			b.setButtonText("Connect and find calendars").onClick(() => void this.discover(b.buttonEl))
		);
		this.statusEl = c.createDiv("pcal-modal-status");
		this.listEl = c.createDiv("pcal-dav-list");
		this.renderList();
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => this.save());
	}

	private normalizedUrl(): string {
		let u = this.serverUrl.trim();
		if (!u) return "";
		if (!/^https?:\/\//i.test(u)) u = "https://" + u;
		return u;
	}

	private async discover(btn: HTMLElement) {
		if (this.busy) return;
		const url = this.normalizedUrl();
		if (!url || !this.username.trim() || !this.password) {
			this.statusEl.setText("Fill in the server, username, and app password first.");
			return;
		}
		this.busy = true;
		btn.setAttribute("disabled", "true");
		this.statusEl.setText("Connecting...");
		try {
			this.found = await discoverCalendars(url, this.username.trim(), this.password);
			const known = new Map(this.collections.map((c) => [c.href, c]));
			this.collections = this.found.map((f, i) => ({
				href: f.href,
				name: f.name,
				color: known.get(f.href)?.color || f.color || paletteColor(i),
				enabled: known.get(f.href)?.enabled ?? true,
			}));
			this.statusEl.setText(`Found ${this.collections.length} calendar${this.collections.length === 1 ? "" : "s"}.`);
			this.renderList();
		} catch (e) {
			this.statusEl.setText(e instanceof CaldavError || e instanceof Error ? e.message : String(e));
		} finally {
			this.busy = false;
			btn.removeAttribute("disabled");
		}
	}

	private renderList() {
		this.listEl.empty();
		for (const coll of this.collections) {
			const row = this.listEl.createDiv("pcal-dav-row");
			const cb = row.createEl("input", { attr: { type: "checkbox" } });
			cb.checked = coll.enabled;
			cb.addEventListener("change", () => (coll.enabled = cb.checked));
			const swatch = row.createSpan("pcal-dav-swatch");
			swatch.style.backgroundColor = coll.color;
			row.createSpan({ text: coll.name });
		}
	}

	private save() {
		const url = this.normalizedUrl();
		if (!url || !this.username.trim()) {
			this.statusEl.setText("A server URL and username are required.");
			return;
		}
		if (!this.collections.length) {
			this.statusEl.setText("Connect and find calendars before saving.");
			return;
		}
		this.onSave({
			id: this.existing?.id ?? freshId(),
			name: this.name.trim() || "CalDAV",
			serverUrl: url,
			username: this.username.trim(),
			password: this.password,
			collections: this.collections,
		});
		this.close();
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Add or edit a vault-notes source: a folder filter plus the property that
 *  holds the date. */
class VaultSourceModal extends Modal {
	private name: string;
	private folder: string;
	private dateProp: string;
	private endProp: string;

	constructor(
		app: App,
		private existing: VaultSource | null,
		private onSave: (src: VaultSource) => void
	) {
		super(app);
		this.name = existing?.name ?? "";
		this.folder = existing?.folder ?? "";
		this.dateProp = existing?.dateProp ?? "date";
		this.endProp = existing?.endProp ?? "";
	}

	onOpen() {
		this.titleEl.setText(this.existing ? "Edit vault source" : "Add vault notes");
		const c = this.contentEl;
		c.createEl("p", {
			cls: "pcal-modal-desc",
			text: "Notes whose frontmatter has the date property appear as events. A bare date (2026-07-17) is all-day; a datetime (2026-07-17T09:30) is timed and can be dragged to reschedule.",
		});
		new Setting(c).setName("Name").addText((t) => t.setPlaceholder("Deadlines").setValue(this.name).onChange((v) => (this.name = v)));
		new Setting(c)
			.setName("Folder")
			.setDesc("Only notes under this folder; empty means the whole vault.")
			.addText((t) => t.setPlaceholder("Projects").setValue(this.folder).onChange((v) => (this.folder = v)));
		new Setting(c).setName("Date property").addText((t) => t.setPlaceholder("date").setValue(this.dateProp).onChange((v) => (this.dateProp = v)));
		new Setting(c)
			.setName("End property")
			.setDesc("Optional: a second property that stretches the event.")
			.addText((t) => t.setPlaceholder("due-end").setValue(this.endProp).onChange((v) => (this.endProp = v)));
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			if (!this.dateProp.trim()) {
				new Notice("Power Desk: a date property name is required.");
				return;
			}
			this.onSave({
				id: this.existing?.id ?? freshId(),
				name: this.name.trim() || "Vault notes",
				folder: this.folder.trim(),
				dateProp: this.dateProp.trim(),
				endProp: this.endProp.trim(),
				color: this.existing?.color ?? "",
				enabled: this.existing?.enabled ?? true,
			});
			this.close();
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

class IcsFeedModal extends Modal {
	private name: string;
	private url: string;

	constructor(
		app: App,
		private existing: IcsFeed | null,
		private onSave: (feed: IcsFeed) => void
	) {
		super(app);
		this.name = existing?.name ?? "";
		this.url = existing?.url ?? "";
	}

	onOpen() {
		this.titleEl.setText(this.existing ? "Edit ICS feed" : "Add ICS feed");
		const c = this.contentEl;
		c.createEl("p", {
			cls: "pcal-modal-desc",
			text: "A read-only iCalendar subscription URL (.ics). Outlook, Google Calendar, and most systems publish one; webcal:// links work too.",
		});
		let urlInput: HTMLInputElement | null = null;
		new Setting(c).setName("Name").addText((t) => t.setPlaceholder("Team calendar").setValue(this.name).onChange((v) => (this.name = v)));
		new Setting(c).setName("Feed URL").addText((t) => {
			urlInput = t.inputEl;
			t.setPlaceholder("https://example.com/calendar.ics").setValue(this.url).onChange((v) => (this.url = v));
		});
		const btns = c.createDiv({ cls: "pcal-modal-btns" });
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			const url = this.url.trim();
			if (!/^(https?|webcal):\/\//i.test(url)) {
				new Notice("Power Desk: that doesn't look like a feed URL.");
				urlInput?.focus();
				return;
			}
			this.onSave({
				id: this.existing?.id ?? freshId(),
				name: this.name.trim() || "ICS feed",
				url,
				color: this.existing?.color ?? "",
				enabled: this.existing?.enabled ?? true,
			});
			this.close();
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

/* ---------------- settings tab ---------------- */

class PCSettingTab extends PluginSettingTab {
	private activeTab = "microsoft";
	private query = "";
	/** Account rows whose calendar lists are expanded; session-scoped. */
	private expandedAccounts = new Set<string>();
	/** The shared-app ID and Tenant fields render only behind the edit pencil. */
	private showGraphAppFields = false;
	private helpEl: HTMLElement | null = null;
	private helpAnchor: HTMLElement | null = null;
	private helpPinned = false;
	private helpCleanup: (() => void) | null = null;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app, plugin);
	}

	hide() {
		this.plugin.refreshSettingsTab = null;
		this.closeHelp();
	}

	private closeHelp() {
		this.helpCleanup?.();
		this.helpCleanup = null;
		this.helpEl?.remove();
		this.helpEl = null;
		this.helpAnchor = null;
		this.helpPinned = false;
	}

	/** Show the help popover for `icon`: a soft theme-colored card rather than
	 *  the native black tooltip. Opens on hover; a click pins it so it survives
	 *  the pointer leaving; Esc, a click elsewhere, or scrolling closes it. */
	private openHelp(icon: HTMLElement, text: string, pin: boolean) {
		if (this.helpAnchor === icon && this.helpEl) {
			if (pin) this.helpPinned = true;
			return;
		}
		this.closeHelp();
		const el = document.body.createDiv({ cls: "pcal-help-pop", text });
		this.helpEl = el;
		this.helpAnchor = icon;
		this.helpPinned = pin;
		const r = icon.getBoundingClientRect();
		el.style.left = Math.max(8, Math.min(r.left - 12, window.innerWidth - el.offsetWidth - 8)) + "px";
		const below = r.bottom + 8;
		el.style.top = (below + el.offsetHeight > window.innerHeight - 8 ? r.top - el.offsetHeight - 8 : below) + "px";
		const onDocDown = (e: MouseEvent) => {
			if (e.target instanceof Node && (el.contains(e.target) || icon.contains(e.target))) return;
			this.closeHelp();
		};
		const onKey = (e: KeyboardEvent) => {
			if (e.key === "Escape") this.closeHelp();
		};
		const onScroll = () => this.closeHelp();
		document.addEventListener("pointerdown", onDocDown, true);
		document.addEventListener("keydown", onKey, true);
		document.addEventListener("scroll", onScroll, true);
		this.helpCleanup = () => {
			document.removeEventListener("pointerdown", onDocDown, true);
			document.removeEventListener("keydown", onKey, true);
			document.removeEventListener("scroll", onScroll, true);
		};
	}

	display() {
		const root = this.containerEl;
		root.empty();
		this.closeHelp(); // a re-render orphans any popover anchored to the old DOM
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		this.plugin.refreshSettingsTab = () => this.display();

		const TABS: { id: string; label: string }[] = [
			{ id: "microsoft", label: "Microsoft 365" },
			{ id: "google", label: "Google" },
			{ id: "caldav", label: "CalDAV" },
			{ id: "ics", label: "ICS feeds" },
			{ id: "vault", label: "Vault notes" },
			{ id: "calendar", label: "Calendar" },
			{ id: "mail", label: "Mail" },
			{ id: "notes", label: "Notes" },
		];
		if (!TABS.some((t) => t.id === this.activeTab)) this.activeTab = TABS[0].id;

		const searchWrap = root.createDiv({ cls: "pcal-settings-search" });
		const searchInput = searchWrap.createEl("input", { cls: "pcal-settings-search-input" });
		searchInput.type = "search";
		searchInput.placeholder = "Search settings...";
		searchInput.value = this.query;

		const tabBar = root.createDiv({ cls: "pcal-settings-tabs" });
		const body = root.createDiv({ cls: "pcal-settings-body" });

		// each heading opens a section div tagged with its tab; the settings that
		// follow render into it because c points at the current section.
		// Add new settings through section(), never a bare setHeading(), or they
		// escape the tabs.
		let c: HTMLElement = body;
		const section = (name: string, tab: string) => {
			const first = !body.querySelector(`.pcal-settings-section[data-tab="${tab}"]`);
			c = body.createDiv({ cls: "pcal-settings-section" });
			c.dataset.tab = tab;
			c.dataset.name = name.toLowerCase();
			if (first) c.dataset.first = "1";
			return new Setting(c).setName(name).setHeading();
		};

		// A section's opening paragraph, built as a real Setting rather than a bare
		// <p>: the theme cards every .setting-item, so loose text floats outside
		// the boxes and breaks the column the rest of the rows line up on.
		const intro = (text: string) => new Setting(c).setDesc(text).setClass("pcal-section-intro");

		// a small help icon after the setting name carrying the deeper "what does
		// this actually do" explanation; hover shows it, a click pins it open. No
		// aria-label on the icon or Obsidian's native black tooltip doubles up.
		const help = (st: Setting, text: string) => {
			const ic = st.nameEl.createSpan({ cls: "pcal-setting-help" });
			setIcon(ic, "help-circle");
			ic.addEventListener("mouseenter", () => this.openHelp(ic, text, false));
			ic.addEventListener("mouseleave", () => {
				if (!this.helpPinned && this.helpAnchor === ic) this.closeHelp();
			});
			ic.addEventListener("click", (e) => {
				e.preventDefault();
				e.stopPropagation();
				if (this.helpPinned && this.helpAnchor === ic) this.closeHelp();
				else this.openHelp(ic, text, true);
			});
		};

		/* ---------------- Sources ---------------- */

		section("Microsoft 365 accounts", "microsoft");
		intro("Work or personal Microsoft accounts, signed in with a device code in your browser; the plugin never sees a password. Each account expands to its calendars, and a friendly name tells work and personal apart at a glance.");
		new Setting(c)
			.setName("Add a Microsoft 365 account")
			.then((st) => help(st, "Starts the sign-in wizard. You sign in on Microsoft's own page in your browser and paste back a short code, so no password ever reaches the plugin. A work account usually rides a shared app registration; a personal one gets its own. Each account you add brings its calendars along and can be renamed afterwards."))
			.setDesc("Work or personal; a short wizard picks the right app registration and signs it in.")
			.addButton((b) => b.setButtonText("Add account").setCta().onClick(() => new GraphAccountWizard(this.app, this.plugin).open()));
		for (const a of s.graphAccounts) {
			const missing: string[] = [];
			if (a.refresh && !this.plugin.canWriteAccount(a)) missing.push("editing");
			if (a.refresh && !this.plugin.canMailAccount(a)) missing.push("mail");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("Mail.ReadWrite")) missing.push("reply windows");
			const open = this.expandedAccounts.has(a.id);
			const row = new Setting(c)
				.setName(this.plugin.nameOf(a))
				.setDesc((a.nickname?.trim() ? a.label + " · " : "") + "Microsoft 365" + (a.clientId.trim() && a.clientId.trim() !== this.plugin.effectiveClientId() ? " · own app" : "") + (!a.refresh ? " · signed out" : missing.length ? ` · reconnect to enable ${missing.join(" and ")}` : ""))
				.setClass("pcal-account-head");
			row.addExtraButton((b) =>
				b.setIcon(open ? "chevron-down" : "chevron-right").setTooltip(open ? "Hide details" : "Rename, inbox, calendars").onClick(() => {
					if (open) this.expandedAccounts.delete(a.id);
					else this.expandedAccounts.add(a.id);
					this.display();
				})
			);
			// any missing power offers the reconnect, not just missing write access
			if (!a.refresh || missing.length) row.addButton((b) => b.setButtonText("Reconnect").setCta().onClick(() => void this.plugin.connectGraph(a)));
			if (a.refresh)
				row.addButton((b) =>
					b.setButtonText("Refresh calendars").onClick(() => {
						void this.plugin.syncGraphCalendars(a).then(() => {
							this.plugin.sourcesChanged();
							this.display();
						});
					})
				);
			row.addButton((b) => b.setButtonText("Remove").setWarning().onClick(() => this.plugin.removeGraphAccount(a)));
			if (!open) continue;
			new Setting(c)
				.setName("Name")
			.then((st) => help(st, "What this account is called throughout the plugin: in the calendar's source list, on mail, and in the account picker. Handy when two accounts share a provider (work and personal Microsoft) and the addresses are the only thing telling them apart. Leaving it empty falls back to the address itself."))
				.setDesc("A friendly name shown wherever this account appears; empty keeps the address.")
				.setClass("pcal-subsetting")
				.addText((t) => {
					t.setPlaceholder(a.label).setValue(a.nickname ?? "").onChange((v) => {
						a.nickname = v;
						save();
					});
					t.inputEl.addEventListener("blur", () => {
						this.plugin.sourcesChanged();
						this.plugin.mailChanged();
						this.display();
					});
				});
			if (this.plugin.canMailAccount(a)) {
				new Setting(c)
					.setName("Inbox in the Mail view")
			.then((st) => help(st, "Whether this account's inbox appears in the Mail view. Turning it off leaves the account's calendars syncing normally and simply stops its mail from showing, which is what you want for an account you read elsewhere."))
					.setClass("pcal-subsetting")
					.addToggle((t) =>
						t.setValue(a.mail !== false).onChange((v) => {
							a.mail = v;
							save();
							this.plugin.mailChanged();
						})
					);
			}
			for (const cal of a.calendars) {
				new Setting(c)
					.setName(cal.name + (cal.isDefault ? " (default)" : ""))
					.setClass("pcal-subsetting")
					.addColorPicker((p) =>
						p.setValue(cal.color || "#888888").onChange((v) => {
							cal.color = v;
							save();
							this.plugin.sourcesChanged();
						})
					)
					.addToggle((t) =>
						t.setValue(cal.enabled).onChange((v) => {
							cal.enabled = v;
							save();
							this.plugin.sourcesChanged();
						})
					);
			}
		}

		section("Microsoft 365 app", "microsoft");
		intro("The shared app registration new work sign-ins ride, borrowed from Power Assistant when present. Add account handles all of this on its own; this section is for inspecting or swapping the shared app. Already-connected accounts keep the app they signed in with.");
		const graphAppReady = !!this.plugin.effectiveClientId();
		const createSt = new Setting(c)
			.setName("Create the app registration")
			.setDesc(
				graphAppReady
					? this.plugin.usingSiblingApp()
						? "Power Assistant's app registration is borrowed automatically; nothing to do here."
						: "A shared app is set; new sign-ins use it."
					: "The Add account wizard creates this inline; these steps set the shared app from here instead."
			);
		help(
			createSt,
			"The registration is a one-time entry in Microsoft's systems that sign-ins name, so consent screens can say who is asking. The guide covers the registration fields, the one switch device-code sign-in needs, and where the Application (client) ID lives: paste that ID below when you have it. When Power Assistant is installed and already set up, its app is borrowed automatically and none of this is needed."
		);
		createSt.addButton((b) => {
			b.setButtonText("Show the steps").onClick(() =>
				new AzureAppGuideModal(this.app, (id) => {
					s.graphClientId = id;
					save();
					this.display();
				}).open()
			);
			if (!graphAppReady) b.setCta();
		});
		const sharedSt = new Setting(c)
			.setName("Shared app")
			.then((st) => help(st, "The Azure app registration that new work sign-ins use. Power Assistant's is borrowed automatically when that plugin is present, so there is usually nothing to set. Accounts already connected keep whichever app they signed in with, so swapping this only affects the next sign-in."))
			.setDesc(
				this.plugin.usingSiblingApp()
					? "Power Assistant's app registration, borrowed automatically."
					: s.graphClientId.trim()
						? s.graphClientId.trim() + (s.graphTenant.trim() ? " · tenant " + s.graphTenant.trim() : "")
						: "None yet; the wizard sets it on the first work sign-in."
			);
		sharedSt.addExtraButton((b) =>
			b.setIcon("pencil").setTooltip(this.showGraphAppFields ? "Hide the fields" : "Edit the ID and tenant").onClick(() => {
				this.showGraphAppFields = !this.showGraphAppFields;
				this.display();
			})
		);
		if (this.showGraphAppFields) {
			const cid = new Setting(c)
				.setName("Application (client) ID")
				.setClass("pcal-subsetting")
				.setDesc(this.plugin.usingSiblingApp() ? "Using Power Assistant's app registration. Enter an ID here to use a different one." : "From your Azure app registration's Overview page.")
				.addText((t) => {
					t.setPlaceholder(this.plugin.usingSiblingApp() ? this.plugin.effectiveClientId() : "00000000-0000-...").setValue(s.graphClientId).onChange((v) => {
						s.graphClientId = v;
						save();
					});
					t.inputEl.addEventListener("blur", () => this.display());
				});
			help(
				cid,
				"From the Overview page of an app created with 'Show the steps' above (any registration with public client flows on works). When Power Assistant is installed and already set up, its app is borrowed automatically and this field can stay empty. Already-connected accounts keep the app they signed in with; this field only steers new sign-ins."
			);
			new Setting(c)
				.setName("Tenant")
			.then((st) => help(st, "Which Microsoft directory the sign-in goes through. 'common' suits personal accounts and most work ones; an organization whose app registration is single-tenant needs its Directory (tenant) ID instead, which an administrator can supply. Getting this wrong shows up as a sign-in that refuses the account rather than anything subtler."))
				.setClass("pcal-subsetting")
				.setDesc("'common' works for most accounts; single-organization apps need their Directory (tenant) ID.")
				.addText((t) => {
					t.setPlaceholder(this.plugin.usingSiblingApp() && this.plugin.effectiveTenant() !== "common" ? this.plugin.effectiveTenant() : "common").setValue(s.graphTenant).onChange((v) => {
						s.graphTenant = v;
						save();
					});
					t.inputEl.addEventListener("blur", () => this.display());
				});
		}

		section("Google accounts", "google");
		intro("Google accounts, signed in through your browser. Each account expands to its calendars, and a friendly name tells accounts apart at a glance.");
		new Setting(c)
			.setName("Add a Google account")
			.then((st) => help(st, "Starts Google's sign-in in your browser. Google requires your own OAuth client (the id and secret above) because installed apps cannot ship a shared one; once that is filled in, adding accounts is a click each."))
			.setDesc(this.plugin.googleReady() ? "Signs in through your browser. Desktop only; the connection syncs." : "Needs your Google Cloud client ID and secret first (Google app, below).")
			.addButton((b) => b.setButtonText("Add account").setCta().onClick(() => void this.plugin.connectGoogle()));
		for (const g of s.googleAccounts) {
			const open = this.expandedAccounts.has(g.id);
			const row = new Setting(c)
				.setName(this.plugin.nameOf(g))
				.setDesc((g.nickname?.trim() ? g.label + " · " : "") + "Google" + (g.refresh ? "" : " · signed out"))
				.setClass("pcal-account-head");
			row.addExtraButton((b) =>
				b.setIcon(open ? "chevron-down" : "chevron-right").setTooltip(open ? "Hide details" : "Rename, calendars").onClick(() => {
					if (open) this.expandedAccounts.delete(g.id);
					else this.expandedAccounts.add(g.id);
					this.display();
				})
			);
			if (!g.refresh) row.addButton((b) => b.setButtonText("Reconnect").setCta().onClick(() => void this.plugin.connectGoogle(g)));
			else
				row.addButton((b) =>
					b.setButtonText("Refresh calendars").onClick(() => {
						void this.plugin.syncGoogleCalendars(g).then(() => {
							this.plugin.sourcesChanged();
							this.display();
						});
					})
				);
			row.addButton((b) => b.setButtonText("Remove").setWarning().onClick(() => this.plugin.removeGoogleAccount(g)));
			if (!open) continue;
			new Setting(c)
				.setName("Name")
			.then((st) => help(st, "What this Google account is called in the calendar's source list and the account picker. Empty keeps the address."))
				.setDesc("A friendly name shown wherever this account appears; empty keeps the address.")
				.setClass("pcal-subsetting")
				.addText((t) => {
					t.setPlaceholder(g.label).setValue(g.nickname ?? "").onChange((v) => {
						g.nickname = v;
						save();
					});
					t.inputEl.addEventListener("blur", () => {
						this.plugin.sourcesChanged();
						this.display();
					});
				});
			for (const cal of g.calendars) {
				new Setting(c)
					.setName(cal.name + (cal.primary ? " (primary)" : "") + (cal.writable ? "" : " · read-only"))
					.setClass("pcal-subsetting")
					.addColorPicker((p) =>
						p.setValue(cal.color || "#888888").onChange((v) => {
							cal.color = v;
							save();
							this.plugin.sourcesChanged();
						})
					)
					.addToggle((t) =>
						t.setValue(cal.enabled).onChange((v) => {
							cal.enabled = v;
							save();
							this.plugin.sourcesChanged();
						})
					);
			}
		}

		section("Google app", "google");
		intro("Google requires your own free Google Cloud project, since its terms forbid shipping shared credentials inside an open plugin. One time: create a project at console.cloud.google.com, enable the Google Calendar API, configure the OAuth consent screen as External and press Publish app, then create an OAuth client of type Desktop app and paste its ID and secret here. The README walks through every step.");
		const gid = new Setting(c)
			.setName("Client ID")
			.addText((t) => t.setPlaceholder("....apps.googleusercontent.com").setValue(s.googleClientId).onChange((v) => {
				s.googleClientId = v;
				save();
			}));
		help(
			gid,
			"Press Publish app on the OAuth consent screen: a project left in Testing mode gets refresh tokens that expire every 7 days, which reads as being signed out weekly. Publishing an External app you never submit for verification just means a one-time 'unverified app' warning at sign-in."
		);
		new Setting(c).setName("Client secret")
			.then((st) => help(st, "The secret half of your Google OAuth client, from the same credentials page as the client id. Google documents the installed-app secret as not confidential, which is why sharing it across your own devices is how they all sign in to the same app. Stored per device and sent only to Google.")).addText((t) => {
			t.inputEl.type = "password";
			t.setPlaceholder("GOCSPX-...").setValue(s.googleClientSecret).onChange((v) => {
				s.googleClientSecret = v;
				save();
			});
		});

		section("CalDAV accounts", "caldav");
		intro("iCloud, Fastmail, Nextcloud, Radicale, and anything else speaking CalDAV. Read-only. Credentials stay in this vault's plugin data; iCloud and Fastmail want an app-specific password, not your account password.");
		for (const account of s.caldavAccounts) {
			new Setting(c)
				.setName(account.name)
				.setDesc(account.serverUrl)
				.addButton((b) =>
					b.setButtonText("Edit").onClick(() => {
						new CaldavAccountModal(this.app, account, (updated) => {
							const i = s.caldavAccounts.findIndex((a) => a.id === updated.id);
							if (i >= 0) s.caldavAccounts[i] = updated;
							s.caldavAccounts = [...s.caldavAccounts];
							save();
							this.plugin.sourcesChanged();
							this.display();
						}).open();
					})
				)
				.addButton((b) =>
					b.setButtonText("Remove").setWarning().onClick(() => {
						s.caldavAccounts = s.caldavAccounts.filter((a) => a.id !== account.id);
						save();
						this.plugin.sourcesChanged();
						this.display();
					})
				);
			for (const coll of account.collections) {
				new Setting(c)
					.setName(coll.name)
					.setClass("pcal-subsetting")
					.addColorPicker((p) =>
						p.setValue(coll.color || "#888888").onChange((v) => {
							coll.color = v;
							save();
							this.plugin.sourcesChanged();
						})
					)
					.addToggle((t) =>
						t.setValue(coll.enabled).onChange((v) => {
							coll.enabled = v;
							save();
							this.plugin.sourcesChanged();
						})
					);
			}
		}
		new Setting(c).setName("Add a CalDAV account")
			.then((st) => help(st, "Connects a calendar server that speaks CalDAV (Fastmail, iCloud, Nextcloud, Radicale, and most self-hosted servers). You provide the server URL and credentials; an app-specific password is the right choice wherever the provider offers one.")).addButton((b) =>
			b.setButtonText("Add account").onClick(() => {
				new CaldavAccountModal(this.app, null, (account) => {
					s.caldavAccounts = [...s.caldavAccounts, account];
					save();
					this.plugin.sourcesChanged();
					this.display();
				}).open();
			})
		);

		section("ICS feeds", "ics");
		intro("Read-only iCalendar subscriptions: holiday calendars, a published Outlook or Google calendar, team schedules.");
		for (const feed of s.icsFeeds) {
			new Setting(c)
				.setName(feed.name)
				.setDesc(feed.url)
				.addColorPicker((p) =>
					p.setValue(feed.color || "#888888").onChange((v) => {
						feed.color = v;
						save();
						this.plugin.sourcesChanged();
					})
				)
				.addToggle((t) =>
					t.setValue(feed.enabled).onChange((v) => {
						feed.enabled = v;
						save();
						this.plugin.sourcesChanged();
					})
				)
				.addButton((b) =>
					b.setButtonText("Edit").onClick(() => {
						new IcsFeedModal(this.app, feed, (updated) => {
							const i = s.icsFeeds.findIndex((f) => f.id === updated.id);
							if (i >= 0) s.icsFeeds[i] = updated;
							s.icsFeeds = [...s.icsFeeds];
							save();
							this.plugin.sourcesChanged();
							this.display();
						}).open();
					})
				)
				.addButton((b) =>
					b.setButtonText("Remove").setWarning().onClick(() => {
						s.icsFeeds = s.icsFeeds.filter((f) => f.id !== feed.id);
						save();
						this.plugin.sourcesChanged();
						this.display();
					})
				);
		}
		new Setting(c).setName("Add an ICS feed")
			.then((st) => help(st, "Subscribes to a read-only calendar published as an .ics URL: a shared team calendar, a sports schedule, a holiday feed. Feeds are fetched on the refresh timer and never written to, so nothing here can change the source.")).addButton((b) =>
			b.setButtonText("Add feed").onClick(() => {
				new IcsFeedModal(this.app, null, (feed) => {
					s.icsFeeds = [...s.icsFeeds, feed];
					save();
					this.plugin.sourcesChanged();
					this.display();
				}).open();
			})
		);

		section("Vault notes", "vault");
		intro("Notes with a date property render as events beside your calendars: deadlines, dated meeting notes, anything. Click one to open the note; drag a timed one to rewrite its date. Entirely local.");
		for (const v of s.vaultSources) {
			new Setting(c)
				.setName(v.name)
				.setDesc(`${v.folder || "Whole vault"} · ${v.dateProp}${v.endProp ? " to " + v.endProp : ""}`)
				.addColorPicker((p) =>
					p.setValue(v.color || "#888888").onChange((val) => {
						v.color = val;
						save();
						this.plugin.sourcesChanged();
					})
				)
				.addToggle((t) =>
					t.setValue(v.enabled).onChange((val) => {
						v.enabled = val;
						save();
						this.plugin.sourcesChanged();
					})
				)
				.addButton((b) =>
					b.setButtonText("Edit").onClick(() => {
						new VaultSourceModal(this.app, v, (updated) => {
							const i = s.vaultSources.findIndex((x) => x.id === updated.id);
							if (i >= 0) s.vaultSources[i] = updated;
							s.vaultSources = [...s.vaultSources];
							save();
							this.plugin.sourcesChanged();
							this.display();
						}).open();
					})
				)
				.addButton((b) =>
					b.setButtonText("Remove").setWarning().onClick(() => {
						s.vaultSources = s.vaultSources.filter((x) => x.id !== v.id);
						save();
						this.plugin.sourcesChanged();
						this.display();
					})
				);
		}
		new Setting(c).setName("Add a vault source")
			.then((st) => help(st, "Turns notes in your vault into calendar events, reading a date property you choose. Useful for anything already tracked as notes (birthdays, deadlines, travel) that you would rather see on the calendar than remember separately.")).addButton((b) =>
			b.setButtonText("Add source").onClick(() => {
				new VaultSourceModal(this.app, null, (src) => {
					s.vaultSources = [...s.vaultSources, src];
					save();
					this.plugin.sourcesChanged();
					this.display();
				}).open();
			})
		);

		/* ---------------- Display ---------------- */

		section("Views", "calendar");
		new Setting(c).setName("Default view")
			.then((st) => help(st, "Which layout the calendar opens on: month for the shape of the weeks ahead, week or work week for hour-by-hour detail, day for one column, agenda for a plain chronological list. Switching views in the calendar itself does not change this, so the view you open on stays predictable.")).addDropdown((d) =>
			d
				.addOptions({ month: "Month", week: "Week", workweek: "Work week", day: "Day", agenda: "Agenda" })
				.setValue(s.defaultMode)
				.onChange((v) => {
					s.defaultMode = v as ViewMode;
					save();
				})
		);
		new Setting(c).setName("Default view on phones")
			.then((st) => help(st, "The view a phone opens on, kept separate because a month grid that reads well on a monitor is cramped on a handset. Agenda is usually the right answer there.")).addDropdown((d) =>
			d
				.addOptions({ agenda: "Agenda", day: "Day", month: "Month", week: "Week", workweek: "Work week" })
				.setValue(s.phoneDefaultMode)
				.onChange((v) => {
					s.phoneDefaultMode = v as ViewMode;
					save();
				})
		);
		new Setting(c).setName("Week starts on")
			.then((st) => help(st, "Which day begins the week in the month grid and week view. Affects only how the calendar is drawn, never the events themselves.")).addDropdown((d) =>
			d
				.addOptions({ monday: "Monday", sunday: "Sunday" })
				.setValue(s.weekStartsMonday ? "monday" : "sunday")
				.onChange((v) => {
					s.weekStartsMonday = v === "monday";
					save();
					this.plugin.notify();
				})
		);
		new Setting(c).setName("24-hour clock")
			.then((st) => help(st, "Show times as 14:00 rather than 2 PM, everywhere the plugin prints a time.")).addToggle((t) =>
			t.setValue(s.use24h).onChange((v) => {
				s.use24h = v;
				save();
				this.plugin.notify();
			})
		);
		new Setting(c).setName("Week numbers")
			.then((st) => help(st, "Show the ISO week number beside each week in the month and week views. Weeks belong to the year containing their Thursday, so the first days of January can carry the previous year's final week number.")).addToggle((t) =>
			t.setValue(s.showWeekNumbers).onChange((v) => {
				s.showWeekNumbers = v;
				save();
				this.plugin.notify();
			})
		);
		const agendaSt = new Setting(c).setName("Agenda covers")
			.then((st) => help(st, "How far ahead the agenda view lists, in days. A longer window means one scroll shows more of what is coming, at the cost of a busier list.")).setDesc(`${s.agendaDays} days`);
		agendaSt.addSlider((sl) =>
			sl
				.setLimits(7, 90, 1)
				.setValue(s.agendaDays)
				.setDynamicTooltip()
				.onChange((v) => {
					s.agendaDays = v;
					agendaSt.setDesc(`${v} days`);
					save();
					this.plugin.notify();
				})
		);
		const dayStartSt = new Setting(c).setName("Scroll the day to").setDesc(`${s.dayStartHour}:00`);
		help(dayStartSt, "Where the week and day grids scroll to when they open. The whole 24 hours stay reachable; this only picks the first hour in view.");
		dayStartSt.addSlider((sl) =>
			sl
				.setLimits(0, 12, 1)
				.setValue(s.dayStartHour)
				.setDynamicTooltip()
				.onChange((v) => {
					s.dayStartHour = v;
					dayStartSt.setDesc(`${v}:00`);
					save();
				})
		);
		const tzSt = new Setting(c).setName("Second time zone").setDesc("Shown beside the hours in week and day views.");
		help(tzSt, "An IANA zone name like Europe/Berlin, Asia/Manila, or America/Los_Angeles. Leave empty for one clock. An unrecognized name simply hides the column rather than erroring.");
		tzSt.addText((t) =>
			t.setPlaceholder("Europe/Berlin").setValue(s.secondTimeZone).onChange((v) => {
				s.secondTimeZone = v;
				save();
				this.plugin.notify();
			})
		);
		const declinedSt = new Setting(c).setName("Show declined events").addToggle((t) =>
			t.setValue(s.showDeclined).onChange((v) => {
				s.showDeclined = v;
				save();
				this.plugin.notify();
			})
		);
		help(declinedSt, "Only Microsoft 365 knows which invites you declined. Off, they disappear; on, they render struck through.");
		const needsSt = new Setting(c).setName("Awaiting response color").addColorPicker((p) =>
			p.setValue(s.calNeedsActionColor).onChange((v) => {
				s.calNeedsActionColor = v;
				save();
				this.plugin.notify();
			})
		);
		help(needsSt, "Background tint for invites you have not accepted yet.");

		section("Refresh", "calendar");
		const remSt = new Setting(c).setName("Meeting reminders").setDesc(s.reminderMinutes > 0 ? `${s.reminderMinutes} minutes before` : "Off");
		help(remSt, "While Obsidian is open, a sticky notice appears before each timed meeting, with a Join button when there is a link. Declined meetings and all-day events stay quiet.");
		remSt.addSlider((sl) =>
			sl
				.setLimits(0, 30, 1)
				.setValue(s.reminderMinutes)
				.setDynamicTooltip()
				.onChange((v) => {
					s.reminderMinutes = v;
					remSt.setDesc(v > 0 ? `${v} minutes before` : "Off");
					save();
				})
		);
		const refreshSt = new Setting(c).setName("Auto-refresh")
			.then((st) => help(st, "How often connected calendars and mail are re-fetched while Obsidian is open. Shorter means fresher and more requests; 0 turns the timer off so nothing is fetched until you refresh by hand.")).setDesc(s.refreshMinutes > 0 ? `Every ${s.refreshMinutes} minute${s.refreshMinutes === 1 ? "" : "s"}` : "Manual only (R or the refresh button)");
		refreshSt.addSlider((sl) =>
			sl
				.setLimits(0, 60, 1)
				.setValue(s.refreshMinutes)
				.setDynamicTooltip()
				.onChange((v) => {
					s.refreshMinutes = v;
					refreshSt.setDesc(v > 0 ? `Every ${v} minute${v === 1 ? "" : "s"}` : "Manual only (R or the refresh button)");
					save();
					this.plugin.refreshCadenceChanged();
				})
		);

		section("Mail", "mail");
		const histSt = new Setting(c)
			.setName("Mail history")
			.then((st) => help(st, "How far back mail is pulled. This is also the ceiling on what Power Assistant's 'Ask your email' can search, because that window only ever indexes messages this plugin has already fetched. Raising it makes the next sync fetch more, once."))
			.setDesc(`Pull the last ${s.mailHistoryDays} days of mail. Also sets how far back Power Assistant's "Ask your email" can reach, since it searches only what is cached here.`);
		histSt.addSlider((sl) =>
			sl
				.setLimits(7, 365, 1)
				.setValue(Math.min(365, Math.max(7, s.mailHistoryDays || 45)))
				.setDynamicTooltip()
				.onChange((v) => {
					s.mailHistoryDays = v;
					histSt.setDesc(`Pull the last ${v} days of mail. Also sets how far back Power Assistant's "Ask your email" can reach, since it searches only what is cached here.`);
					save();
				})
		);
		const capSt = new Setting(c)
			.setName("Messages kept per folder")
			.then((st) => help(st, "How many of the newest messages survive each sync, per folder. This is the setting that actually bounds how much mail you can search: a wide day range changes nothing while this stays low, because older messages are dropped no matter how far back the window reaches. Higher costs memory and a longer first sync."))
			.setDesc(`Retain up to ${s.mailMaxMessages} of the newest messages in each folder. Higher means deeper search and more memory; the initial sync fetches more the first time.`);
		capSt.addSlider((sl) =>
			sl
				.setLimits(50, 5000, 50)
				.setValue(Math.min(5000, Math.max(50, s.mailMaxMessages || 50)))
				.setDynamicTooltip()
				.onChange((v) => {
					s.mailMaxMessages = v;
					capSt.setDesc(`Retain up to ${v} of the newest messages in each folder. Higher means deeper search and more memory; the initial sync fetches more the first time.`);
					save();
				})
		);
		new Setting(c)
			.setName("Saved-mail folder")
			.then((st) => help(st, "Where the Save-to-note button files an email. It has its own setting rather than sharing the calendar's notes folder so that filing mail somewhere specific (a Power Connect encrypted folder, say) does not drag event notes along with it. Empty falls back to the calendar notes folder."))
			.setDesc(
				`Where the "Save to note" button files an email. Empty uses the calendar notes folder (${s.notesFolder.trim() || "Calendar"}). Point it at a Power Connect encrypted folder to keep saved mail encrypted on Dropbox.`
			)
			.addText((t) =>
				t.setPlaceholder(s.notesFolder.trim() || "Calendar").setValue(s.mailNotesFolder).onChange((v) => {
					s.mailNotesFolder = v.trim();
					save();
				})
			);
		new Setting(c).setName("Mark as read")
			.then((st) => help(st, "When a message counts as read. 'As soon as it is selected' matches most mail apps; a short delay avoids marking things read as you arrow past them; 'only when I mark it myself' leaves the decision entirely to you.")).addDropdown((d) =>
			d
				.addOptions({
					select: "As soon as a message is selected",
					delay: "A few seconds after it is selected",
					change: "When the selection changes",
					manual: "Only when I mark it myself",
				})
				.setValue(s.markRead)
				.onChange((v) => {
					s.markRead = v as PCSettings["markRead"];
					save();
					this.display(); // the seconds slider follows the choice
				})
		);
		if (s.markRead === "delay") {
			const secSt = new Setting(c).setName("Mark read after")
			.then((st) => help(st, "How long a message stays selected before it counts as read. Long enough to arrow through a list without clearing everything, short enough that a message you actually stopped on is marked.")).setDesc(`${s.markReadSeconds} second${s.markReadSeconds === 1 ? "" : "s"}`).setClass("pcal-subsetting");
			secSt.addSlider((sl) =>
				sl
					.setLimits(1, 30, 1)
					.setValue(s.markReadSeconds)
					.setDynamicTooltip()
					.onChange((v) => {
						s.markReadSeconds = v;
						secSt.setDesc(`${v} second${v === 1 ? "" : "s"}`);
						save();
					})
			);
		}
		const keepSt = new Setting(c).setName("Unread filter keeps items unread").addToggle((t) =>
			t.setValue(s.unreadFilterKeepsUnread).onChange((v) => {
				s.unreadFilterKeepsUnread = v;
				save();
			})
		);
		help(keepSt, "While the Unread filter is on in the Mail view, selecting a message never marks it read; only the explicit read buttons do. Exactly Outlook's 'always keep items unread' behavior for unread filtering.");
		new Setting(c)
			.setName("Capture orders and bills")
			.then((st) => help(st, "Hands order confirmations and bills to Power Assistant on the refresh timer, which turns them into notes with the vendor, date, and line items filled in. Which messages qualify is decided by rules in Power Assistant's settings, and each match costs one AI call, so this stays off until you have set those rules up."))
			.setDesc(
				"Scan incoming mail on the refresh timer and hand order confirmations and bills to Power Assistant, which turns them into notes you can report on. Needs Power Assistant with a transactions folder set; the rules live in its settings. Each matching message costs one AI call."
			)
			.addToggle((t) =>
				t.setValue(this.plugin.settings.txnScan).onChange((v) => {
					this.plugin.settings.txnScan = v;
					save();
					if (v && !this.plugin.assistantTxn())
						new Notice("Power Desk: Power Assistant is not set up for transactions yet. Set a transactions folder in its settings.", 9000);
				})
			);
		new Setting(c)
			.setName("Save attachments to")
			.then((st) => help(st, "A folder outside the vault that an attachment's 'Save to folder' writes into. This is deliberately not a vault path: it is for files you want on the filesystem rather than in your notes. Empty uses your Downloads folder."))
			.setDesc("The folder outside the vault that an attachment's 'Save to folder' writes into; empty uses your Downloads folder.")
			.addText((t) =>
				t.setValue(s.mailSaveFolder).onChange((v) => {
					s.mailSaveFolder = v;
					save();
				})
			);
		new Setting(c)
			.setName("Signature")
			.then((st) => help(st, "Added to the bottom of new mail, replies, and forwards. Paste HTML to keep formatting, links, and images, or type plain text for something simple. Leave it empty for no signature."))
			.setDesc("Appended to new mail, replies, and forwards. Plain text or HTML both work; pasted HTML keeps its formatting, links, and images.")
			.addTextArea((t) => {
				t.setValue(s.mailSignature).onChange((v) => {
					s.mailSignature = v;
					save();
				});
				t.inputEl.rows = 4;
			});

		section("Availability", "calendar");
		intro("Copy free slots (command palette) reads this window and puts the next five workdays' open times on the clipboard, ready to paste into a mail.");
		const fromSt = new Setting(c).setName("Free day starts")
			.then((st) => help(st, "The earliest hour the free-slot finder will suggest. It never proposes a meeting before this, so an early calendar block does not turn into a breakfast invitation.")).setDesc(`${s.freeFromHour}:00`);
		fromSt.addSlider((sl) =>
			sl
				.setLimits(5, 12, 1)
				.setValue(s.freeFromHour)
				.setDynamicTooltip()
				.onChange((v) => {
					s.freeFromHour = v;
					fromSt.setDesc(`${v}:00`);
					save();
				})
		);
		const toSt = new Setting(c).setName("Free day ends")
			.then((st) => help(st, "The latest hour the free-slot finder will suggest, so a gap late in the evening is not offered as availability.")).setDesc(`${s.freeToHour}:00`);
		toSt.addSlider((sl) =>
			sl
				.setLimits(12, 22, 1)
				.setValue(s.freeToHour)
				.setDynamicTooltip()
				.onChange((v) => {
					s.freeToHour = v;
					toSt.setDesc(`${v}:00`);
					save();
				})
		);

		section("Weather", "calendar");
		intro("The sidebar agenda's day headers carry the forecast from Open-Meteo, a free service needing no account. Type a city and press Look up; clearing the city and looking up again switches weather off.");
		let weatherQuery = s.weatherPlace;
		const locSt = new Setting(c)
			.setName("Location")
			.then((st) => help(st, "The place the forecast is for. Type a town or city and pick from the suggestions; the coordinates are stored, so a renamed or ambiguous place stays resolved to the spot you chose."))
			.setDesc(s.weatherLat.trim() ? `Forecast for ${s.weatherPlace || `${s.weatherLat}, ${s.weatherLon}`}.` : "A city name, like 'Chicago, IL' or 'Fort Wayne'.");
		locSt.addText((t) => t.setValue(s.weatherPlace).onChange((v) => (weatherQuery = v)));
		locSt.addButton((b) =>
			b.setButtonText("Look up").setCta().onClick(async () => {
				const q = weatherQuery.trim();
				if (!q) {
					s.weatherPlace = s.weatherLat = s.weatherLon = "";
					this.plugin.clearWeather();
					save();
					this.display();
					return;
				}
				const hit = await this.plugin.geocodePlace(q);
				if (!hit) {
					new Notice("Power Desk: no place with that name was found.");
					return;
				}
				s.weatherPlace = hit.label;
				s.weatherLat = hit.lat;
				s.weatherLon = hit.lon;
				this.plugin.clearWeather();
				save();
				this.display();
				new Notice(`Power Desk: forecast set to ${hit.label}.`);
			})
		);
		new Setting(c).setName("Unit")
			.then((st) => help(st, "Whether the forecast shows Fahrenheit or Celsius.")).addDropdown((d) =>
			d
				.addOptions({ f: "Fahrenheit", c: "Celsius" })
				.setValue(s.weatherUnit)
				.onChange((v) => {
					s.weatherUnit = v as "f" | "c";
					this.plugin.clearWeather();
					save();
				})
		);

		/* ---------------- Notes ---------------- */

		section("Event notes", "notes");
		intro("Every event can carry a note in your vault: frontmatter for querying, attendees as links so people pages connect, and the body all yours.");
		new Setting(c).setName("Notes folder")
			.then((st) => help(st, "Where an event's linked note is created. Every event can carry one note, and they all land here so the folder doubles as a record of what happened when.")).addText((t) =>
			t.setPlaceholder("Calendar").setValue(s.notesFolder).onChange((v) => {
				s.notesFolder = v;
				save();
				this.plugin.notify();
			})
		);
		const tmpl = new Setting(c).setName("Filename template").addText((t) =>
			t.setPlaceholder("{{date}} {{title}}").setValue(s.noteNameTemplate).onChange((v) => {
				s.noteNameTemplate = v;
				save();
				this.plugin.notify();
			})
		);
		help(tmpl, "Tokens: {{date}} (2026-07-17), {{time}} (09.30, empty for all-day), {{title}}, {{calendar}}. The result is sanitized for the filesystem, so a title's slashes or colons cannot break the path.");
		new Setting(c).setName("Open notes in a new tab")
			.then((st) => help(st, "Whether opening an event's note replaces the current tab or opens beside it. On for keeping the calendar visible while you write.")).addToggle((t) =>
			t.setValue(s.notesInNewTab).onChange((v) => {
				s.notesInNewTab = v;
				save();
			})
		);
		const pplSt = new Setting(c).setName("People folder").setDesc("Where clicking an attendee lands. Empty borrows Power Assistant's People folder.");
		help(pplSt, "Attendee and organizer names on an event card are links: click one to open that person's page, created on first visit. With Power Assistant installed these are the same pages its People hubs build on.");
		pplSt.addText((t) =>
			t.setPlaceholder(this.plugin.personFolderPath()).setValue(s.peopleFolder).onChange((v) => {
				s.peopleFolder = v;
				save();
			})
		);

		/* ---------------- tab bar, search ---------------- */

		const setVisible = (el: HTMLElement, v: boolean) => (el.style.display = v ? "" : "none");
		const applyView = () => {
			const q = this.query.trim().toLowerCase();
			setVisible(tabBar, !q);
			for (const sec of Array.from(body.children) as HTMLElement[]) {
				const items = Array.from(sec.querySelectorAll(":scope > .setting-item:not(.setting-item-heading)")) as HTMLElement[];
				if (!q) {
					for (const it of items) setVisible(it, true);
					setVisible(sec, sec.dataset.tab === this.activeTab);
					continue;
				}
				// a heading-name match reveals the whole section; otherwise match each row
				const nameHit = (sec.dataset.name ?? "").includes(q);
				let anyHit = false;
				for (const it of items) {
					const name = it.querySelector(".setting-item-name")?.textContent?.toLowerCase() ?? "";
					const desc = it.querySelector(".setting-item-description")?.textContent?.toLowerCase() ?? "";
					const hit = nameHit || name.includes(q) || desc.includes(q);
					setVisible(it, hit);
					if (hit) anyHit = true;
				}
				setVisible(sec, anyHit);
			}
		};

		for (const t of TABS) {
			const btn = tabBar.createEl("button", { text: t.label, cls: "pcal-settings-tab" });
			btn.toggleClass("is-active", t.id === this.activeTab);
			btn.onclick = () => {
				if (this.activeTab === t.id) return;
				this.activeTab = t.id;
				for (const other of Array.from(tabBar.children) as HTMLElement[]) other.toggleClass("is-active", other === btn);
				applyView();
			};
		}

		searchInput.addEventListener("input", () => {
			this.query = searchInput.value;
			applyView();
		});

		applyView();
	}
}
