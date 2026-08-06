import { App, ButtonComponent, FuzzySuggestModal, ItemView, MarkdownRenderChild, Menu, Modal, Modifier, Notice, Platform, Plugin, PluginSettingTab, Scope, Setting, SettingDefinitionItem, SettingDefinitionPage, SettingDefinitionRender, SliderComponent, TFile, TFolder, WorkspaceLeaf, arrayBufferToBase64, base64ToArrayBuffer, getIconIds, htmlToMarkdown, normalizePath, requestUrl, sanitizeHTMLToDom, setIcon } from "obsidian";
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
	mailBodyMarkdown,
	mailInlineImages,
	MailEmbed,
	imageExtension,
	imageSize,
	normalizeCid,
	parseDataUrl,
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
	attachmentBadge,
	fmtAttachmentSize,
	mimeForExtension,
	groupThreads,
	MailThread,
	printableHtml,
	printableTableHtml,
	printableAgendaHtml,
	printableMonthHtml,
	PrintDay,
	PrintCell,
	PrintOptions,
	PrintScaleId,
	PRINT_SCALES,
	scaledPt,
	isSystemFolder,
	SignatureUse,
	signatureFor,
	extractInlineImages,
	migrateSignature,
	newArrivals,
	arrivalSummary,
	categoryColor,
	CATEGORY_COLOR_NAMES,
	categoryFolderId,
	folderIdCategory,
	inCategory,
	toggleCategory,
	replaceCategory,
	toGraphDateTime,
	fromGraphDateTime,
	splitSections,
	reconcileChildren,
	MailSection,
	SectionKey,
	UnsubscribeInfo,
	parseUnsubscribe,
	unsubscribePlan,
	RuleEdit,
	EMPTY_RULE,
	ruleToEdit,
	ruleToBody,
	ruleHasUnknownParts,
	ruleSummary,
	MailIndex,
	IndexDoc,
	buildIndex,
	Shortcut,
	groupShortcuts,
	defaultShortcutLabel,
	buildJournal,
	journalMarkdown,
	JournalDay,
	mergePeople,
	SavedContact,
	PersonCard,
	matchEvents,
	eventQueryIsEmpty,
	EventQuery,
	graphSearchText,
	passesLocalFilters,
	parseQuery,
	buildQuery,
	SearchFields,
	searchIndex,
	whenPresets,
	fmtWhen,
	ContactHit,
	rankContacts,
	matchContacts,
	currentAddressFragment,
	applyAddressChoice,
	GeoHit,
	pickGeoHit,
	splitPlaceQuery,
	stripMeetingBoilerplate,
	weatherGlyph,
	RepeatKind,
	sanitizeName,
	snapMin,
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
	fetchUnreadInFolder,
	getEvent,
	getInboxId,
	getMailAttachmentBytes,
	getMessage,
	listMailAttachments,
	getSchedule,
	listCalendars,
	addFileAttachment,
	addInlineImage,
	Receipts,
	OutgoingFile,
	postOneClickUnsubscribe,
	getUserPhoto,
	AutoReply,
	getAutoReply,
	setAutoReply,
	OutlookCategory,
	listContacts,
	TodoList,
	TodoTask,
	StickyNote,
	listUnreadIdsInFolder,
	patchReadBatch,
	findNotesFolder,
	listStickyNotes,
	getStickyNoteBody,
	createStickyNote,
	listTodoLists,
	listTodoTasks,
	createTodoTask,
	setTodoTaskDone,
	deleteTodoTask,
	listMasterCategories,
	fetchMessagesByCategory,
	setMessageCategories,
	createCategory,
	updateCategoryColor,
	deleteCategory,
	findMessagesByCategory,
	patchCategoriesBatch,
	markAsJunk,
	reportMessageBeta,
	GraphRule,
	listMessageRules,
	createMessageRule,
	updateMessageRule,
	deleteMessageRule,
	createDraftMessage,
	ensureMailFolder,
	createMailFolder,
	renameMailFolder,
	deleteMailFolder,
	flagMessage,
	markMessageRead,
	moveMessage,
	setDeferredSend,
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
/** How many inline pictures one saved message may bring into the vault. A
 *  newsletter can carry dozens of them; a note wants the ones that are the
 *  message, not every badge in its footer. */
const MAX_NOTE_IMAGES = 20;

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
	/** Show this account's inbox in the mail view; undefined means yes. */
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
	/** Whether a saved email brings its inline pictures into the vault. Off,
	 *  the note keeps the words and the pictures are left in the mailbox. */
	mailNoteImages: boolean;
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
	/** How many lines of body preview sit under each message row, exactly
	 *  Outlook's View > Message Preview. Defaults to 0, which is Outlook
	 *  Classic's own default: sender over subject, nothing else, which fits
	 *  half again as many messages on a screen and is the fastest to scan
	 *  once the mail is familiar. */
	mailPreviewLines: 0 | 1 | 2 | 3;
	/** Group the list into conversations, Outlook's Show as Conversations and
	 *  Gmail's threads. Search results are never grouped: a search asks about
	 *  messages, and hiding matches inside a collapsed thread would lie about
	 *  what matched. */
	mailConversations: boolean;
	/** How much air each row gets. Separate from the preview lines on
	 *  purpose: one decides how much a row says, this decides how much room
	 *  it says it in. */
	mailDensity: "compact" | "cozy" | "comfortable";
	/** Messages parked in the Snoozed folder and when they are due back.
	 *  `messageId` is the id in the Snoozed folder, since a move mints a new
	 *  one, and `returnFolderId` is where it came from. Lives in settings so
	 *  it survives a restart and follows the vault to another machine. */
	mailSnoozes: { accountId: string; messageId: string; dueMs: number; returnFolderId: string; subject: string }[];
	/** The folder snoozed mail waits in, created in the mailbox on demand. */
	mailSnoozeFolder: string;
	/** Seconds a sent message is held so it can be taken back; 0 turns it
	 *  off and mail goes the instant you press Send. */
	mailUndoSeconds: number;
	/** How often, in seconds, an open mail view checks for new mail. Separate
	 *  from the calendar's cadence because the two are not comparable: a
	 *  meeting moving is rare and a message arriving is constant, and mail
	 *  syncs by delta, so a check that finds nothing costs one small request.
	 *  0 leaves mail to the calendar's timer and the Refresh button. */
	mailPollSeconds: number;
	/** Whether new mail announces itself, and how choosily. "focused" follows
	 *  Outlook's own verdict and stays quiet about the Other pile. */
	mailNotify: "off" | "focused" | "all";
	/** An unread count on the mail ribbon icon. */
	mailBadge: boolean;
	/** The toolbar's buttons, by id, in the order they appear. Anything left
	 *  out is still reachable by shortcut, right-click, and the palette. */
	mailToolbar: string[];
	/** The calendar header's action buttons, by id and in order. The nav, the
	 *  title, and the view switcher are not in here: those are the view
	 *  rather than actions on it, and a calendar you cannot page or switch is
	 *  not a calendar. */
	calendarToolbar: string[];
	/** Toolbar additions that have been offered once, by key.
	 *
	 *  An action that ships after someone has already arranged their toolbar
	 *  is invisible until they go looking for it, which is the difference
	 *  between shipping a feature and hiding one. Each new action is appended
	 *  to an existing toolbar exactly once; take it off afterwards and it
	 *  stays off, because this never runs for that key again. */
	toolbarAdded: string[];
	/** Show a sender's real profile picture where their initials would be.
	 *  Needs the photo permission, which is a reconnect away. */
	mailPhotos: boolean;
	/** Leave focus mode automatically when you move to a note, so mail can be
	 *  full width and the vault can be itself without a toggle between them. */
	/** Split the inbox into Priority, Focused, Notifications, and Other,
	 *  the way Spark and Shortwave bundle a list. Applies to the triage
	 *  surfaces only: the unified list, an inbox, and Unread Mail. Browsing
	 *  Sent Items or a filing folder stays flat, and a search always does. */
	mailSplitInbox: boolean;
	/** Folders filed into lately, newest first, so the move picker leads with
	 *  the handful anyone actually uses instead of a whole tree. */
	mailRecentFolders: { accountId: string; folderId: string; name: string }[];
	/** Folders dragged into an order of their own, per account. Only the ones
	 *  actually moved are listed; everything else stays alphabetical behind
	 *  them, so arranging three folders does not freeze the other thirty. */
	mailFolderOrder: { accountId: string; folderIds: string[] }[];
	/** Sections folded shut, by key. */
	mailSectionsCollapsed: string[];
	/** Scan incoming mail for orders and bills and hand matches to Power
	 *  Assistant. Off by default: it spends an AI call per matched message. */
	txnScan: boolean;
	/** The one signature this plugin used to have. Kept only so an existing
	 *  one can be carried into the list below on first run. */
	mailSignature: string;
	/** Named signatures, as Outlook has them. */
	/** Whether a new message asks for a read or delivery receipt by default.
	 *  Both off, because asking every correspondent to confirm they read you
	 *  is a thing to choose per message, not a habit to acquire silently. */
	mailAskReadReceipt: boolean;
	mailAskDeliveryReceipt: boolean;
	/** Outlook's Shortcuts, which here can point at a mail folder, a saved
	 *  search, a note, or a link, since this plugin straddles both halves. */
	shortcuts: Shortcut[];
	mailSignatures: { id: string; name: string; html: string }[];
	/** Which signature each account uses, separately for new mail and for
	 *  replies, since the long one belongs on a first message and rarely on
	 *  the fourth reply of a thread. An empty id means none. */
	mailSignatureUse: { accountId: string; newId: string; replyId: string }[];
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
	// deep enough that scrolling back and searching find something, which is
	// the whole reason to hold mail locally. The list draws row by row rather
	// than all at once, so a long list costs no more to keep current than a
	// short one; what it costs is the first sync and the memory.
	mailMaxMessages: 500,
	agendaDays: 30,
	notesFolder: "Calendar",
	mailNotesFolder: "",
	mailNoteImages: true,
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
	mailPreviewLines: 0,
	mailConversations: true,
	mailDensity: "cozy",
	mailSnoozes: [],
	mailSnoozeFolder: "Snoozed",
	mailUndoSeconds: 8,
	mailPollSeconds: 30,
	mailNotify: "focused",
	mailBadge: true,
	// Outlook's own order, which is what the hands already know
	mailToolbar: ["delete", "report", "reply", "replyAll", "forward", "readUnread", "categorize", "flag", "snooze", "move", "archive", "makeEvent", "saveToNote", "print"],
	calendarToolbar: ["filter", "refresh", "findEvent", "freeSlots", "print"],
	toolbarAdded: [],
	mailPhotos: true,
	mailSplitInbox: true,
	mailSectionsCollapsed: [],
	mailRecentFolders: [],
	mailFolderOrder: [],
	txnScan: false,
	mailSignature: "",
	mailAskReadReceipt: false,
	mailAskDeliveryReceipt: false,
	shortcuts: [],
	mailSignatures: [],
	mailSignatureUse: [],
	peopleFolder: "",
	phoneDefaultMode: "agenda",
};

const freshId = () => Math.random().toString(36).slice(2, 10);

/** The first of these icon names Obsidian actually ships, falling back to the
 *  last one given.
 *
 *  setIcon takes any string and quietly draws nothing for a name it does not
 *  know, and IconName is typed as plain string, so neither the compiler nor
 *  the runtime complains about a guess: the button simply comes out empty.
 *  Asking which ids exist costs one call and makes the fallback deliberate.
 *  Ids arrive bare or lucide-prefixed depending on the version, so both
 *  spellings count as a match. */
let iconIds: Set<string> | null = null;
function pickIcon(...names: string[]): string {
	if (!iconIds) {
		try {
			iconIds = new Set(getIconIds());
		} catch {
			iconIds = new Set();
		}
	}
	for (const n of names) if (iconIds.has(n) || iconIds.has(`lucide-${n}`)) return n;
	return names[names.length - 1];
}

/** Paint a button as destructive.
 *
 *  `setDestructive` arrived in 1.13 and this plugin's floor is 1.8.7, where
 *  calling it would throw, so the old `setWarning` has to stay reachable. The
 *  cast is the runtime check: the inline type carries no deprecation, which is
 *  also what keeps the fallback from being reported as one. */
function markDestructive(b: ButtonComponent): ButtonComponent {
	const btn = b as unknown as { setDestructive?: () => void; setWarning: () => void };
	if (btn.setDestructive) btn.setDestructive();
	else btn.setWarning();
	return b;
}

/** Keep a slider's value visible while it is dragged.
 *
 *  1.13 shows it inline and retired `setDynamicTooltip`, but on 1.8.7 the call
 *  is the only thing that shows the number at all, so it is reached through a
 *  cast rather than named: absent on new builds, harmless on old ones. */
function showSliderValue(sl: SliderComponent): SliderComponent {
	(sl as unknown as { setDynamicTooltip?: () => void }).setDynamicTooltip?.();
	return sl;
}

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

/** The forecast, by day key. Named because an empty one has to be created in
 *  two places, and `new Map()` on its own infers nothing. */
type WeatherDays = Map<string, { hi: number; lo: number; code: number }>;

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
	/** data.json's size and mtime as we last saw them, so the desktop poll can
	 *  tell a file someone else wrote from one nobody touched, without reading
	 *  it; see watchDataFile(). */
	private dataStamp: string | null = null;
	private autoTimer: number | null = null;

	async onload() {
		this.adoptSettings(Object.assign({}, DEFAULT_SETTINGS, (await this.loadData()) as Partial<PCSettings> | null));
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
		this.migrateSignatures();
		this.offerToolbarActions();
		await this.loadCacheFile();
		const stale = this.settings.graphAccounts.filter((a) => a.refresh && a.grantedScope !== GRAPH_SCOPE);
		if (stale.length) {
			new Notice(`Power Desk: reconnect ${stale.map((a) => this.nameOf(a)).join(", ")} in settings to enable the newest permissions (event editing, mail, reply windows).`, 10000);
		}

		this.registerView(VIEW_TYPE, (leaf) => new PowerCalendarView(leaf, this));
		this.registerView(VIEW_TYPE_MAIL, (leaf) => new MailView(leaf, this));
		this.addRibbonIcon("calendar-days", "Open Power Desk", () => void this.openCalendarView());
		this.ribbonEl = this.addRibbonIcon("mail", "Open Power Desk inbox", () => void this.openMailView());
		this.paintRibbonBadge();
		this.addCommand({ id: "open-mail", icon: "inbox", name: "Open inbox", callback: () => void this.openMailView() });
		this.addCommand({ id: "new-mail", icon: "pencil", name: "New mail", callback: () => new RichComposeModal(this.app, this, { mode: "new" }).open() });
		// the palette and the shortcut card are worth having on the vault's
		// own command list too, so they can be given a global hotkey rather
		// than only working once the mail view has focus
		this.addCommand({
			id: "mail-palette",
			icon: "search",
			name: "Mail actions",
			callback: () => {
				void this.openMailView().then(() => {
					const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
					if (v instanceof MailView) new MailPaletteModal(this.app, this, v).open();
				});
			},
		});
		this.addCommand({ id: "mail-shortcuts", icon: "keyboard", name: "Mail keyboard shortcuts", callback: () => new MailShortcutsModal(this.app).open() });
		this.addCommand({ id: "mail-rules", icon: "filter", name: "Inbox rules", callback: () => new RulesModal(this.app, this).open() });
		this.addCommand({ id: "people", icon: "users", name: "People", callback: () => this.openPeople() });
		this.addCommand({ id: "tasks", icon: "check-square", name: "Tasks", callback: () => this.openTasks() });
		this.addCommand({ id: "notes", icon: "sticky-note", name: "Outlook notes", callback: () => new NotesModal(this.app, this).open() });
		this.addCommand({ id: "journal", icon: "book-open", name: "Journal (a day of mail and meetings)", callback: () => new JournalModal(this.app, this).open() });
		this.addCommand({ id: "folders", icon: "folder", name: "Folders", callback: () => this.openFolders() });
		this.addCommand({ id: "shortcuts", icon: "bookmark", name: "Shortcuts", callback: () => this.openShortcuts() });
		this.addCommand({
			id: "print-calendar",
			icon: "printer",
			name: "Print the calendar",
			callback: () => {
				const v = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0]?.view;
				if (v instanceof PowerCalendarView) v.printCalendar();
				else new Notice("Power Desk: open the calendar first, so there is a range to print.");
			},
		});
		this.addCommand({ id: "mail-ooo", icon: "megaphone", name: "Automatic replies (out of office)", callback: () => new OutOfOfficeModal(this.app, this).open() });
		this.addCommand({
			id: "focus-mode",
			icon: "maximize",
			name: "Focus mode (full width mail and calendar)",
			callback: () => {
				// entering from a note would focus on nothing, so bring mail up
				if (!this.focusOn() && !this.onPowerView()) void this.openMailView().then(() => this.toggleFocus(true));
				else this.toggleFocus();
			},
		});
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
		this.addCommand({
			id: "search-events",
			icon: "search",
			name: "Find events (by title, people, place, or date)",
			callback: () => {
				void this.openCalendarView().then((v) => v?.openEventSearch());
			},
		});
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
		// snoozed mail comes back on its own clock rather than off the mail
		// refresh, which would recurse: waking a message refreshes the mail
		this.registerInterval(window.setInterval(() => void this.wakeSnoozes(), 60_000));
		this.scheduleMailPoll();
		this.register(() => {
			if (this.mailPollTimer != null) window.clearInterval(this.mailPollTimer);
		});
		// coming back to the window, or to the mail view, is exactly when
		// stale mail is most obvious, so both check straight away rather than
		// waiting out whatever is left of the current interval
		this.registerDomEvent(window, "focus", () => this.pollMailNow());
		this.registerEvent(
			this.app.workspace.on("active-leaf-change", (leaf) => {
				const v = leaf?.view;
				if (v instanceof MailView) this.pollMailNow();
				// The full-width fold follows the tab it belongs to. Opening
				// mail or the calendar still rearranges nothing on its own —
				// that is still the panel button's job alone — but once that
				// button has been pressed, the fold is part of what those two
				// tabs look like, and stepping onto a note hands the pages
				// back rather than leaving them shut behind it.
				this.applyFocus();
			})
		);
		// anything on the ribbon opens something in the page tree, so bring
		// the tree back rather than firing a click into a collapsed pane
		this.registerDomEvent(
			document,
			"click",
			(e) => {
				if (!this.pagesFolded) return;
				const t = e.target as HTMLElement | null;
				// the tree comes back for the click, but the fold stays armed:
				// coming back to mail should not need the button pressed again
				if (t?.closest?.(".workspace-ribbon")) this.unfoldNow();
			},
			true
		);
		// never leave the tree collapsed behind us
		this.register(() => this.toggleFocus(false));
		// a message still inside its undo window when the plugin unloads has
		// not been sent yet, so let it go rather than losing it
		this.register(() => this.flushPendingSends());
		this.app.workspace.onLayoutReady(() => void this.wakeSnoozes());
		// know about a running out-of-office from the first paint, so the
		// banner is there rather than appearing a beat later
		this.app.workspace.onLayoutReady(() => {
			for (const a of this.mailAccounts()) void this.loadAutoReply(a.id);
		});
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
	previewReminder() {
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
	 *  paragraph, because the three things it has to answer, which meeting, how
	 *  long have I got, how do I get in, are answered at three different
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
		// The card is styled as `.notice.pcal-remind-notice`, so the classes go on
		// the notice's own box rather than the message inside it. Walking up from
		// messageEl finds that box on every build, and closest() counts the element
		// itself, so it is right either way round.
		notice.messageEl.closest(".notice")?.addClass("pcal-remind-notice", "pw-self-styled");
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

	private dataPath(): string {
		return `${this.app.vault.configDir}/plugins/${this.manifest.id}/data.json`;
	}

	/**
	 * Desktop: notice that someone else rewrote data.json, so external edits are
	 * adopted while you are looking rather than at the next restart.
	 *
	 * onExternalSettingsChange covers Obsidian's own Sync. A folder sync
	 * (Dropbox, OneDrive, Drive) landing another device's settings is a plainer
	 * event than that, and can arrive unannounced, so the file's own size and
	 * mtime are the signal. Asking the vault adapter for them keeps this inside
	 * the vault: it stats one known file under the config folder, never reaches
	 * for a path of its own, and reads nothing until the stamp actually moves.
	 * Our own saves move it too, and cost one wasted read that adoptExternalData
	 * recognizes as its own echo and drops.
	 */
	private watchDataFile() {
		if (!Platform.isDesktopApp) return;
		this.registerInterval(window.setInterval(() => void this.checkDataFile(), 5000));
		// coming back to the window is when another device's change is most
		// likely to be sitting there waiting, so look straight away
		this.registerDomEvent(window, "focus", () => void this.checkDataFile());
	}

	private async checkDataFile() {
		if (this.busySaving()) return; // our own write is on its way; let it land
		try {
			const st = await this.app.vault.adapter.stat(this.dataPath());
			if (!st) return;
			const stamp = `${st.mtime}:${st.size}`;
			const first = this.dataStamp === null;
			if (stamp === this.dataStamp) return;
			this.dataStamp = stamp;
			if (!first) await this.adoptExternalData();
		} catch {
			/* unreadable this moment (a sync mid-swap); the next tick tries again */
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
				photos?: Record<string, string>;
			};
			if (c.v !== 1 && c.v !== 2) return;
			for (const [k, st] of Object.entries(c.events ?? {})) {
				if (!k.startsWith("vault:")) this.cache.set(k, { ...st, error: null, inFlight: false });
			}
			for (const [k, st] of Object.entries(c.mail ?? {})) this.mailCache.set(k, { deltaLink: null, ...st, error: null, inFlight: false });
			for (const [k, st] of Object.entries(c.folders ?? {})) this.mailFolderCache.set(k, { ...st, inFlight: false });
			for (const [k, st] of Object.entries(c.folderLists ?? {})) this.folderCache.set(k, { deltaLink: null, ...st, error: null, inFlight: false });
			for (const [k, b] of Object.entries(c.bodies ?? {})) this.bodyCache.set(k, b);
			for (const [k, p] of Object.entries(c.photos ?? {})) this.photoCache.set(k, p);
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
			// faces ride along so they are there on the next launch rather
			// than fading in one request at a time; only the ones that exist,
			// and capped, since these are the only bytes here worth counting
			const photos: Record<string, string> = {};
			let n = 0;
			for (const [k, url] of this.photoCache) {
				if (!url || n >= 400) continue;
				photos[k] = url;
				n++;
			}
			await this.app.vault.adapter.write(this.cachePath(), JSON.stringify({ v: 2, events, mail, folders, folderLists, bodies, photos }));
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

	private weatherCache: { fetchedAt: number; days: WeatherDays } | null = null;
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
			const days: WeatherDays = new Map();
			(d.daily?.time ?? []).forEach((t, i) => {
				days.set(t, { hi: Math.round(d.daily?.temperature_2m_max?.[i] ?? 0), lo: Math.round(d.daily?.temperature_2m_min?.[i] ?? 0), code: d.daily?.weather_code?.[i] ?? 0 });
			});
			this.weatherCache = { fetchedAt: Date.now(), days };
			this.notify();
		} catch {
			// a failed fetch keeps what it had and retries in an hour
			this.weatherCache = { fetchedAt: Date.now(), days: this.weatherCache?.days ?? new Map<string, { hi: number; lo: number; code: number }>() };
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
		this.searchIdxDirty = true;
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
			const fm: Record<string, unknown> | undefined = this.app.metadataCache.getFileCache(f)?.frontmatter;
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

	private mailPollTimer: number | null = null;

	/** How stale an open mail list is allowed to get. The mail cadence when
	 *  one is set, otherwise the calendar's. */
	mailStaleMs(): number {
		const secs = this.settings.mailPollSeconds;
		return secs > 0 ? secs * 1000 : Math.max(1, this.settings.refreshMinutes) * 60000;
	}

	/** Check for new mail, but only what is actually being looked at: the
	 *  per-account inbox lists behind the unified view, the folder tree that
	 *  carries the unread counts, and whichever folder each open view is
	 *  showing.
	 *
	 *  Deliberately not every folder ever visited. Those stay cached and
	 *  refresh when opened, because polling thirty of them a minute would
	 *  spend a great deal of somebody's API quota redrawing nothing. */
	private pollMail() {
		const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL);
		if (!leaves.length) return;
		this.ensureMail(true);
		this.ensureMailFolders(true);
		for (const leaf of leaves) {
			const v = leaf.view;
			if (v instanceof MailView) v.pollCurrentFolder();
		}
	}

	private scheduleMailPoll() {
		if (this.mailPollTimer != null) window.clearInterval(this.mailPollTimer);
		this.mailPollTimer = null;
		if (this.settings.mailPollSeconds > 0) this.mailPollTimer = window.setInterval(() => this.pollMail(), Math.max(15, this.settings.mailPollSeconds) * 1000);
	}

	private lastPollAt = 0;

	/** A poll now rather than on the next tick: coming back to the window or
	 *  to the mail view is exactly when stale mail is most obvious, and it is
	 *  also when a timer is least likely to have just fired.
	 *
	 *  Throttled, because window focus and active-leaf changes arrive in
	 *  bursts: collapsing a sidebar can fire several between one blink and
	 *  the next, and a check a second after the last one has nothing new to
	 *  find while costing a round trip per account and a repaint. */
	pollMailNow() {
		const now = Date.now();
		if (now - this.lastPollAt < 3000) return;
		this.lastPollAt = now;
		this.pollMail();
	}

	mailCadenceChanged() {
		this.scheduleMailPoll();
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
	 *  `canceled` is consulted between polls. Resolves to the connected
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
				return googleToPC(raw, { sourceId: def.key, calendarName: def.label, color: def.color, writable: def.writable });
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
			new Notice(ev.recurring ? "Power Desk: occurrence canceled." : "Power Desk: event deleted.");
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
		const staleMs = this.mailStaleMs();
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
			this.announceArrivals(a.id, st.messages);
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
		// the shape of the tree changes rarely, but its unread counts change
		// every time mail arrives, which is the whole point of polling
		const staleMs = this.mailStaleMs();
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
			// the tree lists categories beside the folders, so they load with
			// it rather than waiting for the first Categorize menu. One read
			// per session per account, and none without the scope.
			void this.loadCategories(a.id);
			void this.prefetchFolders(a);
		}
	}

	/** Whether a folder is an account's inbox, which decides whether the
	 *  split inbox applies to it. */
	isInboxFolder(accountId: string, folderId: string): boolean {
		return this.mailFolderCache.get(accountId)?.inboxId === folderId;
	}

	folderTreeFor(a: GraphAccount): { folder: MailFolder; depth: number; expandable: boolean }[] {
		const st = this.mailFolderCache.get(a.id);
		return st ? orderFolderTree(st.folders, st.inboxId, new Set(this.settings.mailCollapsed), this.folderOrderFor(a.id)) : [];
	}

	folderOrderFor(accountId: string): string[] {
		return this.settings.mailFolderOrder.find((o) => o.accountId === accountId)?.folderIds ?? [];
	}

	/** Put one folder where another sits, among their shared siblings.
	 *
	 *  The stored order only ever names the siblings of the pair being
	 *  arranged, seeded from how they read right now, so dragging inside one
	 *  branch leaves every other branch alphabetical and free to take new
	 *  folders as they appear. */
	reorderFolder(accountId: string, dragId: string, targetId: string) {
		if (dragId === targetId) return;
		const a = this.accountById(accountId);
		const st = this.mailFolderCache.get(accountId);
		if (!a || !st) return;
		const parentOf = (id: string) => st.folders.find((f) => f.id === id)?.parentId ?? null;
		// only siblings rearrange: moving a folder under a different parent is
		// a mailbox change, not a display order
		if (parentOf(dragId) !== parentOf(targetId)) return;
		const siblings = this.folderTreeFor(a)
			.map((n) => n.folder)
			.filter((f) => f.id === dragId || parentOf(f.id) === parentOf(targetId));
		const ids = siblings.map((f) => f.id);
		const from = ids.indexOf(dragId);
		const to = ids.indexOf(targetId);
		if (from < 0 || to < 0) return;
		ids.splice(to, 0, ids.splice(from, 1)[0]);
		// keep every other branch's arrangement, replace this one's
		const kept = this.folderOrderFor(accountId).filter((id) => !ids.includes(id));
		const next = [...kept, ...ids];
		this.settings.mailFolderOrder = [...this.settings.mailFolderOrder.filter((o) => o.accountId !== accountId), { accountId, folderIds: next }];
		this.queueSave();
		this.notify();
	}

	/** Give one account's folders back their alphabetical order. */
	resetFolderOrder(accountId: string) {
		this.settings.mailFolderOrder = this.settings.mailFolderOrder.filter((o) => o.accountId !== accountId);
		this.queueSave();
		this.notify();
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
	private mailboxGates = new Map<string, { active: number; queue: { wake: () => void; idle: boolean }[] }>();

	/** How badly this request is wanted.
	 *
	 *  "now" is someone looking at a spinner: the reading pane, an attachment
	 *  being opened. "idle" is prefetch, which exists to make "now" rare and
	 *  must never be the reason "now" waits. */
	async mailboxGate<T>(accountId: string, fn: () => Promise<T>, prio: "now" | "soon" | "idle" = "soon"): Promise<T> {
		const g = this.mailboxGates.get(accountId) ?? { active: 0, queue: [] };
		this.mailboxGates.set(accountId, g);
		// Prefetch gets one slot of the three, so a click always has two free
		// to start in. Without this the reading pane could sit behind a batch
		// of twenty bodies, which is exactly the two or three seconds it took
		// to open a message while the list was warming.
		const limit = prio === "idle" ? 1 : 3;
		while (g.active >= limit) {
			await new Promise<void>((res) => {
				const w = { wake: res, idle: prio === "idle" };
				if (prio === "now") g.queue.unshift(w);
				else g.queue.push(w);
			});
		}
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
			// wake the first waiter that can actually start: waking a prefetch
			// that must go straight back to sleep would swallow the wake-up and
			// leave everyone behind it waiting on the next request to finish
			const i = g.queue.findIndex((w) => g.active < (w.idle ? 1 : 3));
			if (i >= 0) g.queue.splice(i, 1)[0].wake();
		}
	}

	ensureFolderMail(accountId: string, folderId: string, force = false) {
		const a = this.accountById(accountId);
		if (!a) return;
		const key = `${accountId}:${folderId}`;
		const st = this.folderCache.get(key);
		if (st?.inFlight) return;
		const staleMs = this.mailStaleMs();
		if (!force && st && Date.now() - st.fetchedAt < staleMs) return;
		void this.fetchFolderMailFor(a, folderId);
	}

	/** `quiet` is the prefetcher warming a folder nobody has opened: it queues
	 *  behind anything a person is waiting on, and it does not announce itself,
	 *  since a list you cannot see changing is not news. */
	private async fetchFolderMailFor(a: GraphAccount, folderId: string, quiet = false) {
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
		if (!quiet) this.notify();
		try {
			await this.mailboxGate(
				a.id,
				async () => {
					const category = folderIdCategory(folderId);
					if (category) {
						const cap = Math.min(2000, Math.max(50, this.settings.mailMaxMessages || 500));
						const raw = await fetchMessagesByCategory(await this.graphTokenFor(a), category, cap);
						const fresh = raw.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label)).filter((m): m is PCMail => m != null);
						st.messages = fresh.sort((x, y) => y.receivedMs - x.receivedMs).slice(0, cap);
					} else if (folderId === UNREAD_FOLDER) {
						const fresh = await this.unreadAcrossInbox(a);
						// This folder is a question the mailbox answers afresh every
						// time, so a message you have just read is simply not in the
						// answer. Replacing the list wholesale would delete that row
						// a moment after you marked it, leaving nothing to mark
						// unread again; the ones you touched yourself keep their
						// place until a refresh you actually asked for.
						const ids = new Set(fresh.map((m) => m.id));
						const kept = st.messages.filter((m) => this.recentlyMarked.has(m.id) && !ids.has(m.id));
						// a mark still in flight means the mailbox can answer with the
						// old state; what you did wins over what it has not heard yet,
						// so a row does not go bold again for a moment
						for (const m of fresh) {
							const read = this.recentlyMarked.get(m.id);
							if (read !== undefined) m.unread = !read;
						}
						st.messages = [...fresh, ...kept].sort((x, y) => y.receivedMs - x.receivedMs);
					} else {
						await this.syncList(a, folderId, st);
					}
				},
				quiet ? "idle" : "soon"
			);
			st.error = null;
		} catch (e) {
			st.error = e instanceof Error ? e.message : String(e);
		} finally {
			st.fetchedAt = Date.now();
			st.inFlight = false;
			// a warmed folder nobody has opened has nothing to say; the folder
			// pane's counts come from the tree, not from these lists
			if (!quiet) this.notify();
			this.queueCachePersist();
		}
	}

	/** Every unread message in the inbox and everything under it.
	 *
	 *  A folder at a time, and only the folders the tree says are holding
	 *  something. The mailbox-wide alternative this replaced asked for the
	 *  newest hundred unread ANYWHERE and then threw away whatever fell
	 *  outside the inbox, which is fine until the mailbox has thousands of
	 *  unread somewhere else: one real mailbox spent 81 of that hundred on
	 *  sync-conflict copies and another on junk, and showed 18 of its 33.
	 *  Worse, the ones it dropped were the OLDEST, so a quiet folder's unread
	 *  was invisible however long it sat there.
	 *
	 *  Reading the same counts the folder pane reads is the point: the list
	 *  and the number beside it now come from one answer, so they cannot
	 *  disagree. It costs one request per folder that has unread in it, which
	 *  is a handful, and none at all for a mailbox that is caught up. */
	private async unreadAcrossInbox(a: GraphAccount): Promise<PCMail[]> {
		const token = await this.graphTokenFor(a);
		const toPC = (raw: unknown[]) => raw.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label)).filter((m): m is PCMail => m != null);
		const st = this.mailFolderCache.get(a.id);
		// before the tree has landed there is nothing to walk, so the old
		// mailbox-wide sweep stands in for one round
		if (!st || !st.inboxId || !st.folders.length) return toPC(await fetchUnreadMessages(token));
		const subtree = folderSubtreeIds(st.folders, st.inboxId);
		const picks = st.folders.filter((f) => subtree.has(f.id) && f.unread > 0);
		const cap = Math.min(5000, Math.max(50, this.settings.mailMaxMessages || 500));
		const out: PCMail[] = [];
		// already inside the mailbox gate, so these go one after another
		// rather than taking slots a click is waiting for
		for (const f of picks) {
			if (out.length >= cap) break;
			out.push(...toPC(await fetchUnreadInFolder(token, f.id, Math.min(cap, f.unread + 10))));
		}
		return out.sort((x, y) => y.receivedMs - x.receivedMs).slice(0, cap);
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
				await this.fetchFolderMailFor(a, f.id, true);
			}
		} finally {
			this.prefetchingFolders.delete(a.id);
		}
	}

	private bodyPrefetchRunning = false;
	private bodyPrefetchNext: PCMail[] | null = null;

	/** Read the top of a list before it is clicked, so the reading pane is
	 *  instant for anything visible. Silent; a miss costs nothing.
	 *
	 *  A list that arrives mid-run is remembered rather than dropped: mail
	 *  polls every thirty seconds, and simply returning meant the newest
	 *  arrivals, the ones actually about to be clicked, were the ones least
	 *  likely to be warm. */
	async prefetchBodies(list: PCMail[]) {
		if (this.bodyPrefetchRunning) {
			this.bodyPrefetchNext = list;
			return;
		}
		this.bodyPrefetchRunning = true;
		try {
			let next: PCMail[] | null = list;
			while (next) {
				const run = next;
				this.bodyPrefetchNext = null;
				await this.readMailBodies(run.slice(0, 40), true);
				next = this.bodyPrefetchNext;
			}
		} finally {
			this.bodyPrefetchRunning = false;
			this.bodyPrefetchNext = null;
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
					// idle: a batch of twenty is the slowest thing this plugin
					// asks for, and nobody is watching it
					out = await this.mailboxGate(accountId, () => getMessagesBatch(token, group), quiet ? "idle" : "soon");
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
			// A prefetched body is invisible by definition: nothing in the list
			// draws body text, and the reading pane awaits its own message's
			// body rather than waiting to be told. Notifying here redrew the
			// whole view twice per poll for a result nobody could see.
			if (!quiet) this.notify();
		}
		return cached;
	}

	/** Turn one Graph message into a cached body entry. Shared by the single
	 *  and batched readers so both store exactly the same shape. */
	private cacheMailBody(m: PCMail, raw: Record<string, unknown>): { text: string; html?: string; toLine: string; unsub?: UnsubscribeInfo } {
		const body = raw.body as { contentType?: string; content?: string } | undefined;
		const content = body?.content ?? "";
		const isHtml = body?.contentType?.toLowerCase() === "html" || /<\w+[^>]*>/.test(content);
		const text = isHtml ? stripHtml(content) : content;
		const toLine = ((raw.toRecipients as { emailAddress?: { name?: string; address?: string } }[] | undefined) ?? [])
			.map((r) => r.emailAddress?.name || r.emailAddress?.address || "")
			.filter(Boolean)
			.join(", ");
		const unsub = parseUnsubscribe(raw.internetMessageHeaders as { name?: string; value?: string }[] | undefined) ?? undefined;
		const result = { text, html: isHtml ? content : "", toLine, unsub };
		this.bodyCache.set(m.id, result);
		if (this.bodyCache.size > 150) {
			const oldest = this.bodyCache.keys().next().value;
			if (oldest) this.bodyCache.delete(oldest);
		}
		return result;
	}

	folderMail(accountId: string, folderId: string): PCMail[] {
		const messages = this.folderCache.get(`${accountId}:${folderId}`)?.messages ?? [];
		// taking the category off a message is done from inside this very list,
		// so the row has to leave as soon as it does rather than at the next
		// fetch: the cached message object is the one Categorize just rewrote
		const category = folderIdCategory(folderId);
		if (category) return messages.filter((m) => inCategory(m, category));
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

	/** How much mail a category holds.
	 *
	 *  This counts everything, not the unread, which is where it parts company
	 *  with the folders above it in the tree. A folder's number answers how much
	 *  is waiting for you; a category's answers how much carries the label, and
	 *  that is the number the list shows when you click it. Counting unread put
	 *  a 2 beside a list of 3 and made the row look broken, which it was: a
	 *  category's count and its own list must never disagree.
	 *
	 *  It counts what has been fetched rather than asking the mailbox, which has
	 *  no per-category total to give and would want a query per category on
	 *  every draw of the pane. So a category nobody has opened shows no count
	 *  until it has been, and one larger than the fetch cap counts to the cap,
	 *  exactly as its list does. */
	categoryCount(accountId: string, name: string): number {
		return this.folderMail(accountId, categoryFolderId(name)).length;
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

	private mailSearch: { query: string; results: PCMail[]; inFlight: boolean; error: string | null; scope: "local" | "everywhere" } | null = null;

	mailSearchState(): { query: string; results: PCMail[]; inFlight: boolean; error: string | null; scope: "local" | "everywhere" } | null {
		return this.mailSearch;
	}

	/* ----- the local index ----- */

	private searchIdx: MailIndex | null = null;
	/** Set whenever anything about the cached mail may have moved. The index
	 *  is rebuilt on the next search rather than here, so a refresh that
	 *  nobody is searching over costs nothing. */
	private searchIdxDirty = true;

	markSearchIndexStale() {
		this.searchIdxDirty = true;
	}

	/** The index over everything cached: every message in every list, with
	 *  the body of the ones whose body has been read.
	 *
	 *  This is the honest shape of a local index in a plugin. It cannot know
	 *  about mail that has never been fetched, and it can only search the
	 *  bodies it has actually seen, which is why the search box says what it
	 *  covered rather than implying it looked everywhere. What it buys is the
	 *  thing a server round trip cannot: results that narrow on every
	 *  keystroke. */
	private ensureSearchIndex(): MailIndex {
		if (this.searchIdx && !this.searchIdxDirty) return this.searchIdx;
		const docs: IndexDoc[] = [];
		const seen = new Set<string>();
		for (const a of this.mailAccounts()) {
			for (const list of this.cachedMailLists(a.id)) {
				for (const m of list.messages) {
					if (seen.has(m.id)) continue;
					seen.add(m.id);
					docs.push({
						id: m.id,
						subject: m.subject,
						from: `${m.from} ${m.fromAddress}`,
						// the preview is always there; the full body only once read
						body: `${m.preview} ${this.bodyCache.get(m.id)?.text ?? ""}`,
						ms: m.receivedMs,
						unread: m.unread,
						flagged: !!m.flagged,
						hasAttachments: !!m.hasAttachments,
					});
				}
			}
		}
		this.searchIdx = buildIndex(docs);
		this.searchIdxDirty = false;
		return this.searchIdx;
	}

	/** Every cached message by id, for turning index hits back into mail. */
	private allCachedById(): Map<string, PCMail> {
		const out = new Map<string, PCMail>();
		for (const a of this.mailAccounts()) for (const list of this.cachedMailLists(a.id)) for (const m of list.messages) if (!out.has(m.id)) out.set(m.id, m);
		return out;
	}

	/** Instant results from the index, with no network at all. Runs on every
	 *  keystroke, which is why it has to stay local. */
	searchLocal(query: string): void {
		const q = query.trim();
		if (!q) {
			this.clearMailSearch();
			return;
		}
		const byId = this.allCachedById();
		const hits = searchIndex(this.ensureSearchIndex(), q, 200)
			.map((id) => byId.get(id))
			.filter((m): m is PCMail => !!m);
		this.mailSearch = { query: q, results: hits, inFlight: false, error: null, scope: "local" };
		this.notify();
	}

	/** Search every mail account's whole mailbox; one account failing costs
	 *  its results, not the search. A newer query supersedes a slower one. */
	async runMailSearch(query: string): Promise<void> {
		const q = query.trim();
		if (!q) {
			this.clearMailSearch();
			return;
		}
		// the local hits go up first so there is something to read while the
		// mailbox is asked, then the server's answer merges into them
		const byId = this.allCachedById();
		const local = searchIndex(this.ensureSearchIndex(), q, 200)
			.map((id) => byId.get(id))
			.filter((m): m is PCMail => !!m);
		const state = { query: q, results: local, inFlight: true, error: null as string | null, scope: "local" as "local" | "everywhere" };
		this.mailSearch = state;
		this.notify();
		// the mailbox is only asked what it can answer; the flags and the date
		// bounds are applied here, to its results as well as ours
		const parsed = parseQuery(q);
		const serverText = graphSearchText(parsed);
		const all: PCMail[] = [];
		let firstError: string | null = null;
		if (serverText.trim()) {
			for (const a of this.mailAccounts()) {
				try {
					const raw = await searchMessages(await this.graphTokenFor(a), serverText, 25);
					all.push(
						...raw
							.map((m) => graphMailToPC(m as GraphMailLike, a.id, this.nameOf(a), a.label))
							.filter((m): m is PCMail => m != null)
							.filter((m) => passesLocalFilters(m, parsed))
					);
				} catch (e) {
					firstError = firstError ?? (e instanceof Error ? e.message : String(e));
				}
			}
		}
		if (this.mailSearch !== state) return;
		// the server is authoritative for what it returns, but it caps at 25
		// per account and does not know the local query language, so the
		// local hits it did not name are kept rather than thrown away
		const merged = new Map<string, PCMail>();
		for (const m of all) merged.set(m.id, m);
		for (const m of local) if (!merged.has(m.id)) merged.set(m.id, m);
		state.results = [...merged.values()].sort((x, y) => y.receivedMs - x.receivedMs);
		state.error = firstError;
		state.inFlight = false;
		state.scope = "everywhere";
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

	/** Refresh because someone asked for one, rather than because a timer
	 *  fired. That difference matters: this is the moment the lists are
	 *  allowed to forget what you marked and show what the mailbox says. */
	userRefreshMail() {
		this.recentlyMarked.clear();
		this.refreshMailAll(true);
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
	private bodyCache = new Map<string, { text: string; html?: string; toLine: string; unsub?: UnsubscribeInfo }>();

	/** The full body for the reading pane: HTML mail keeps its HTML for the
	 *  sanitized rich rendering, with the stripped text riding along for notes
	 *  and previews. Quiet mode is for prefetch, where a failure should cost
	 *  nothing visible. */
	/** The body if it is already here, without a promise.
	 *
	 *  Awaiting a cache hit still costs a turn of the event loop and a second
	 *  paint, which is the difference between a message that appears with the
	 *  click and one that appears just after it. */
	cachedMailBody(m: PCMail): { text: string; html?: string; toLine: string; unsub?: UnsubscribeInfo } | null {
		// entries cached before HTML rendering existed lack the html field
		// entirely (even text mail now stores html: ""); those want a refetch
		const hit = this.bodyCache.get(m.id);
		return hit && hit.html !== undefined ? hit : null;
	}

	/** Fetch a body nobody is waiting for yet, once. Hovering a row and moving
	 *  on must not queue the same request five times. */
	private warmingBodies = new Set<string>();

	warmBody(m: PCMail) {
		if (this.warmingBodies.has(m.id) || this.cachedMailBody(m)) return;
		this.warmingBodies.add(m.id);
		void this.readMailBody(m, true, "idle").finally(() => this.warmingBodies.delete(m.id));
	}

	async readMailBody(m: PCMail, quiet = false, prio: "now" | "soon" | "idle" = "now"): Promise<{ text: string; html?: string; toLine: string; unsub?: UnsubscribeInfo } | null> {
		const hit = this.cachedMailBody(m);
		if (hit) return hit;
		const a = this.accountById(m.accountId);
		if (!a) return null;
		try {
			const raw = await this.mailboxGate(a.id, async () => getMessage(await this.graphTokenFor(a), m.id), prio);
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

	/** Messages whose read state you changed yourself.
	 *
	 *  An unread-only list would otherwise drop a message the instant you
	 *  marked it read, which is the one moment you might want to change your
	 *  mind about it. These keep their place until the list is genuinely
	 *  reloaded. */
	readonly recentlyMarked = new Map<string, boolean>();

	async setMailRead(m: PCMail, read: boolean) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		this.recentlyMarked.set(m.id, read);
		// bounded, because nothing clears it on a session where you never
		// refresh and never change folder
		if (this.recentlyMarked.size > 300) this.recentlyMarked.delete(this.recentlyMarked.keys().next().value as string);
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

	/** Flag or clear a message for follow-up. The list updates first and the
	 *  mailbox catches up, same as marking read: a flag that waited on a
	 *  round trip would feel broken on a keypress. */
	async setMailFlag(m: PCMail, flagged: boolean) {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			const cached = list.messages.find((x) => x.id === m.id);
			if (cached) cached.flagged = flagged || undefined;
		}
		this.notify();
		try {
			await flagMessage(await this.graphTokenFor(a), m.id, flagged);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/* ----- undo send ----- */

	/** Sends being held for their undo window, by id. */
	private pendingSends = new Map<string, { timer: number; fire: () => Promise<void> }>();

	/** Hold a send briefly so it can be taken back, then let it go.
	 *
	 *  The window closes the moment you hit Send, because the point is that
	 *  sending feels done; the message simply has not left yet. Undo puts the
	 *  window back exactly as it was, draft and all.
	 *
	 *  The honest limit: this is a client-side hold, so a message still in
	 *  its undo window when Obsidian quits has not been sent. Unloading the
	 *  plugin flushes anything waiting rather than dropping it, which covers
	 *  a normal quit or a plugin reload; a force quit or a power cut is
	 *  beyond reach, which is why the window is seconds rather than minutes. */
	holdSend(opts: { label: string; fire: () => Promise<void>; reopen: () => void }): void {
		const secs = Math.max(0, this.settings.mailUndoSeconds);
		if (!secs) {
			void opts.fire();
			return;
		}
		const id = freshId();
		const timer = window.setTimeout(() => {
			this.pendingSends.delete(id);
			void opts.fire();
		}, secs * 1000);
		this.pendingSends.set(id, { timer, fire: opts.fire });

		const frag = createFragment();
		const wrap = frag.createDiv("pcal-undo-notice");
		wrap.createSpan({ text: `Sending "${opts.label}"` });
		const btn = wrap.createEl("button", { text: "Undo" });
		const notice = new Notice(frag, secs * 1000);
		btn.addEventListener("click", () => {
			const p = this.pendingSends.get(id);
			if (!p) return;
			window.clearTimeout(p.timer);
			this.pendingSends.delete(id);
			notice.hide();
			opts.reopen();
		});
	}

	/** Let every held send go now. Unload calls this so quitting Obsidian
	 *  during an undo window still puts the mail on its way. */
	flushPendingSends() {
		for (const [id, p] of this.pendingSends) {
			window.clearTimeout(p.timer);
			this.pendingSends.delete(id);
			void p.fire();
		}
	}

	openShortcuts() {
		new ShortcutsModal(
			this.app,
			this,
			(accountId, folderId, name) => {
				void this.openMailView().then(() => {
					const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
					if (v instanceof MailView) v.goFolder(accountId, folderId, name);
				});
			},
			(q) => {
				void this.openMailView().then(() => {
					const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
					if (v instanceof MailView) v.runSearch(q, true);
				});
			}
		).open();
	}

	openFolders() {
		new FoldersModal(this.app, this, (accountId, folderId, name) => {
			void this.openMailView().then(() => {
				const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
				if (v instanceof MailView) v.goFolder(accountId, folderId, name);
			});
		}).open();
	}

	openTasks() {
		new TasksModal(this.app, this, (m) => {
			void this.openMailView().then(() => {
				const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
				if (v instanceof MailView) v.revealMessage(m);
			});
		}).open();
	}

	/** The People window, wired so its actions land somewhere: a mail search
	 *  in the inbox, a meeting search in the calendar. Both views are opened
	 *  if they are not already, since an action that silently does nothing
	 *  because the wrong tab was in front is not an action. */
	openPeople() {
		new PeopleModal(
			this.app,
			this,
			(q) => {
				void this.openMailView().then(() => {
					const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
					if (v instanceof MailView) v.runSearch(q, true);
				});
			},
			(name) => {
				void this.openCalendarView().then((v) => v?.openEventSearchFor(name));
			}
		).open();
	}

	/* ----- recipient autocomplete ----- */

	/** Sightings harvested from Sent Items, kept for the session. The cached
	 *  mail and the calendar are read fresh on every keystroke instead,
	 *  since they are already in memory and cost nothing. */
	private sentSeen: { name?: string; email: string; ms: number }[] = [];
	private sentLoaded = false;

	/** Who to suggest in a recipient box.
	 *
	 *  Deliberately built from what this plugin can already see rather than
	 *  from Graph's People API: /me/people needs the People.Read scope, and
	 *  adding a scope makes every connected account reconnect before anything
	 *  works again. That is a steep toll for a convenience. Three sources
	 *  cover it well without asking for anything new: who you have written to
	 *  (Sent Items, the strongest signal by far), who has written to you
	 *  (every cached list), and who you have meetings with (the calendar). */
	contactIndex(): ContactHit[] {
		const seen: { name?: string; email: string; ms: number }[] = [...this.sentSeen];
		for (const a of this.mailAccounts()) {
			for (const list of this.cachedMailLists(a.id)) {
				for (const m of list.messages) if (m.fromAddress) seen.push({ name: m.from, email: m.fromAddress, ms: m.receivedMs });
			}
		}
		for (const st of this.cache.values()) {
			for (const ev of st.events) {
				for (const p of ev.attendeeDetail ?? []) if (p.email) seen.push({ name: p.name, email: p.email, ms: ev.startMs });
			}
		}
		// a saved contact belongs in autocomplete even if you have never
		// written to them, which is most of the point of saving one
		for (const c of this.savedContactList) seen.push({ name: c.name, email: c.email, ms: 0 });
		// never suggest the mailboxes doing the sending
		const own = new Set(this.settings.graphAccounts.map((a) => (a.label ?? "").toLowerCase()).filter(Boolean));
		return rankContacts(seen).filter((c) => !own.has(c.email));
	}

	/** The mailbox's saved contacts, read once a session. */
	private savedContactList: SavedContact[] = [];
	private contactsLoaded = false;
	private contactsDenied = false;

	savedContacts(): SavedContact[] {
		return this.savedContactList;
	}

	/** Whether any mail account could read contacts if asked. */
	contactsNeedReconnect(): boolean {
		const accounts = this.mailAccounts();
		return accounts.length > 0 && !accounts.some((a) => a.grantedScope.includes("Contacts"));
	}

	async ensureSavedContacts(): Promise<void> {
		if (this.contactsLoaded || this.contactsDenied) return;
		this.contactsLoaded = true;
		const able = this.mailAccounts().filter((a) => a.grantedScope.includes("Contacts"));
		if (!able.length) return;
		const all: SavedContact[] = [];
		for (const a of able) {
			try {
				all.push(...(await listContacts(await this.graphTokenFor(a))));
			} catch (e) {
				// a refusal means the permission is not really there; stop
				// rather than failing once per account on every open
				if (/403|forbidden|scope|permission/i.test(e instanceof Error ? e.message : String(e))) this.contactsDenied = true;
			}
		}
		this.savedContactList = all;
		this.notify();
	}

	/** The address book and the correspondence as one list. */
	people(): PersonCard[] {
		return mergePeople(this.contactIndex(), this.savedContactList);
	}

	/** Read Sent Items once a session and remember who was on the To lines.
	 *  This is what makes the first few suggestions the right ones: mail you
	 *  received tells you who talks at you, Sent Items tells you who you
	 *  actually write to. */
	async ensureSentContacts(): Promise<void> {
		if (this.sentLoaded) return;
		this.sentLoaded = true;
		for (const a of this.mailAccounts()) {
			try {
				const raw = await fetchFolderMessages(await this.graphTokenFor(a), "sentitems", 100);
				for (const r of raw) {
					const o = r as { receivedDateTime?: string; toRecipients?: { emailAddress?: { name?: string; address?: string } }[]; ccRecipients?: { emailAddress?: { name?: string; address?: string } }[] };
					const ms = Date.parse(o.receivedDateTime ?? "");
					if (!Number.isFinite(ms)) continue;
					for (const p of [...(o.toRecipients ?? []), ...(o.ccRecipients ?? [])]) {
						const email = p.emailAddress?.address;
						if (email) this.sentSeen.push({ name: p.emailAddress?.name, email, ms });
					}
				}
			} catch {
				// no Sent Items, no permission, no network: the cached mail and
				// the calendar still make a usable list, so this stays quiet
			}
		}
		this.notify();
	}

	/* ----- snooze ----- */

	/** Park a message until a time: into the mailbox's Snoozed folder, with
	 *  a note of where it came from so it can go back there.
	 *
	 *  There is no server-side snooze in Graph, so the return leg is ours to
	 *  run and can only run while Obsidian is open. That is a real limit and
	 *  the UI says so rather than pretending otherwise: a message due at 8am
	 *  comes back the first time the app is running after 8am. Parking it in
	 *  a real Outlook folder rather than hiding it in the plugin means the
	 *  worst case is a visible folder of mail you can file by hand, not mail
	 *  that has vanished. */
	async snoozeMail(m: PCMail, dueMs: number): Promise<boolean> {
		const a = this.accountById(m.accountId);
		if (!a) return false;
		try {
			const token = await this.graphTokenFor(a);
			const folderId = await ensureMailFolder(token, this.settings.mailSnoozeFolder.trim() || "Snoozed");
			if (!folderId) return false;
			const returnFolderId = m.folderId || (await getInboxId(token)) || "inbox";
			const newId = await moveMessage(token, m.id, folderId);
			this.settings.mailSnoozes = [
				...this.settings.mailSnoozes.filter((s) => s.messageId !== m.id),
				{ accountId: a.id, messageId: newId ?? m.id, dueMs, returnFolderId, subject: m.subject },
			];
			this.queueSave();
			// the message is somewhere else now, so drop it from every cached
			// list rather than leaving a row that points at a dead id
			for (const list of this.cachedMailLists(a.id)) list.messages = list.messages.filter((x) => x.id !== m.id);
			this.notify();
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** Anything due goes back where it came from, unread so it reads as new
	 *  arrival. Runs at load and on every mail refresh. */
	async wakeSnoozes(): Promise<void> {
		const now = Date.now();
		const due = this.settings.mailSnoozes.filter((s) => s.dueMs <= now);
		if (!due.length) return;
		const woke: string[] = [];
		for (const s of due) {
			const a = this.accountById(s.accountId);
			// an account that has gone away takes its snoozes with it, else
			// they would sit in the list forever being retried
			if (!a) {
				this.settings.mailSnoozes = this.settings.mailSnoozes.filter((x) => x !== s);
				continue;
			}
			try {
				const token = await this.graphTokenFor(a);
				const newId = await moveMessage(token, s.messageId, s.returnFolderId);
				await markMessageRead(token, newId ?? s.messageId, false);
				this.settings.mailSnoozes = this.settings.mailSnoozes.filter((x) => x !== s);
				woke.push(s.subject);
			} catch (e) {
				// a message deleted or moved by hand while it was snoozed will
				// never come back; drop it rather than retrying every refresh
				const msg = e instanceof Error ? e.message : String(e);
				if (/not ?found|ErrorItemNotFound|404/i.test(msg)) this.settings.mailSnoozes = this.settings.mailSnoozes.filter((x) => x !== s);
			}
		}
		if (woke.length) {
			this.queueSave();
			new Notice(woke.length === 1 ? `Power Desk: "${woke[0]}" is back in your inbox.` : `Power Desk: ${woke.length} snoozed messages are back in your inbox.`);
			this.refreshMailAll(true);
		}
	}

	/** What is waiting, soonest first, for the snoozed list in the view. */
	snoozedFor(accountId?: string) {
		return this.settings.mailSnoozes.filter((s) => !accountId || s.accountId === accountId).sort((a, b) => a.dueMs - b.dueMs);
	}

	/** Bring one back early, from the snoozed list. */
	async unsnooze(rec: { accountId: string; messageId: string; returnFolderId: string }): Promise<void> {
		const a = this.accountById(rec.accountId);
		if (!a) return;
		try {
			const token = await this.graphTokenFor(a);
			const newId = await moveMessage(token, rec.messageId, rec.returnFolderId);
			await markMessageRead(token, newId ?? rec.messageId, false);
		} catch (e) {
			this.graphErrorNotice(e);
		}
		this.settings.mailSnoozes = this.settings.mailSnoozes.filter((x) => x.messageId !== rec.messageId);
		this.queueSave();
		this.refreshMailAll(true);
	}

	/* ----- inbox rules ----- */

	/** Rules run in the mailbox, not here, which is the whole point of them:
	 *  mail is filed while Obsidian is closed, on the phone, everywhere. This
	 *  plugin only edits them. */
	private rulesCache = new Map<string, GraphRule[]>();
	/** Why an account's rules could not be read, if they could not. Kept
	 *  apart from the rules themselves so an empty list and a failed read are
	 *  never shown as the same thing: "no rules yet" over a permission error
	 *  is a lie, and a confident one. */
	private rulesError = new Map<string, string>();

	rulesFor(accountId: string): GraphRule[] {
		return this.rulesCache.get(accountId) ?? [];
	}

	rulesErrorFor(accountId: string): string | null {
		return this.rulesError.get(accountId) ?? null;
	}

	/** Whether an account was connected before rules needed their own
	 *  permission. Checked up front so the answer is "reconnect this account"
	 *  rather than a bare 403 from a request that could not have worked. */
	rulesNeedReconnect(accountId: string): boolean {
		const a = this.accountById(accountId);
		return !!a?.refresh && !a.grantedScope.includes("MailboxSettings");
	}

	async loadRules(accountId: string): Promise<GraphRule[]> {
		const a = this.accountById(accountId);
		if (!a) return [];
		if (this.rulesNeedReconnect(accountId)) {
			this.rulesError.set(accountId, `Reconnect ${this.nameOf(a)} in settings to let Power Desk read its rules. Rules sit under a separate mailbox permission that this account has not granted yet.`);
			this.notify();
			return [];
		}
		try {
			const rules = await listMessageRules(await this.graphTokenFor(a));
			this.rulesCache.set(accountId, rules);
			this.rulesError.delete(accountId);
			this.notify();
			return rules;
		} catch (e) {
			const msg = e instanceof Error ? e.message : String(e);
			this.rulesError.set(
				accountId,
				/403|forbidden|scope|permission/i.test(msg)
					? `${this.nameOf(a)} is not allowed to read its rules yet. Add the delegated MailboxSettings.ReadWrite permission to the Azure app registration, then Reconnect the account in settings. (${msg})`
					: msg
			);
			this.notify();
			return [];
		}
	}

	/** Create or update a rule. An existing rule is handed to the merge so
	 *  whatever Outlook set that this editor does not show comes back out the
	 *  other side intact. */
	async saveRule(accountId: string, edit: RuleEdit, existing?: GraphRule): Promise<boolean> {
		const a = this.accountById(accountId);
		if (!a) return false;
		try {
			const token = await this.graphTokenFor(a);
			const body = ruleToBody(edit, existing);
			if (existing) await updateMessageRule(token, existing.id, body);
			else await createMessageRule(token, { ...body, sequence: (this.rulesFor(accountId).length || 0) + 1 });
			await this.loadRules(accountId);
			new Notice(existing ? "Power Desk: rule updated." : "Power Desk: rule created. It runs in your mailbox from now on.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	async removeRule(accountId: string, ruleId: string): Promise<void> {
		const a = this.accountById(accountId);
		if (!a) return;
		try {
			await deleteMessageRule(await this.graphTokenFor(a), ruleId);
			await this.loadRules(accountId);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/** Report as junk: blocks the sender and files it in Junk Email, which is
	 *  what trains the filter. The supported v1.0 action. */
	async reportJunk(targets: PCMail[]): Promise<number> {
		let n = 0;
		for (const m of targets) {
			const a = this.accountById(m.accountId);
			if (!a) continue;
			for (const list of this.cachedMailLists(a.id)) list.messages = list.messages.filter((x) => x.id !== m.id);
			try {
				await markAsJunk(await this.graphTokenFor(a), m.id);
				n++;
			} catch (e) {
				this.graphErrorNotice(e);
			}
		}
		this.notify();
		return n;
	}

	/** Report phishing, or rescue something wrongly filed as junk. Both run
	 *  on the preview endpoint, so a failure is reported as itself rather
	 *  than quietly turned into some other action the user did not ask for. */
	async reportPreview(targets: PCMail[], action: "phish" | "notJunk"): Promise<{ done: number; error: string | null }> {
		let done = 0;
		let error: string | null = null;
		for (const m of targets) {
			const a = this.accountById(m.accountId);
			if (!a) continue;
			try {
				await reportMessageBeta(await this.graphTokenFor(a), m.id, action, true);
				for (const list of this.cachedMailLists(a.id)) list.messages = list.messages.filter((x) => x.id !== m.id);
				done++;
			} catch (e) {
				error = error ?? (e instanceof Error ? e.message : String(e));
			}
		}
		this.notify();
		return { done, error };
	}

	/* ----- new mail notifications ----- */

	/** The newest message each account had when this session started. Set on
	 *  the first sync rather than persisted: reopening Obsidian after a week
	 *  should not replay a week of arrivals, and mail that came in while it
	 *  was closed is not news either. */
	private notifyBaseline = new Map<string, number>();
	private told = new Set<string>();
	private ribbonEl: HTMLElement | null = null;

	/** Announce whatever just landed, and keep the ribbon's count honest. */
	private announceArrivals(accountId: string, messages: PCMail[]) {
		const mode = this.settings.mailNotify;
		const base = this.notifyBaseline.get(accountId);
		if (base === undefined) {
			// first sync of the session: take the high-water mark and say nothing
			this.notifyBaseline.set(accountId, messages.reduce((n, m) => Math.max(n, m.receivedMs), 0));
			this.paintRibbonBadge();
			return;
		}
		if (mode === "off") {
			this.notifyBaseline.set(accountId, messages.reduce((n, m) => Math.max(n, m.receivedMs), base));
			this.paintRibbonBadge();
			return;
		}
		const arrivals = newArrivals(messages, base, this.told, mode === "focused");
		this.notifyBaseline.set(accountId, messages.reduce((n, m) => Math.max(n, m.receivedMs), base));
		this.paintRibbonBadge();
		if (!arrivals.length) return;
		for (const a of arrivals) this.told.add(a.id);
		if (this.told.size > 500) this.told = new Set([...this.told].slice(-300));

		const { title, body } = arrivalSummary(arrivals);
		const frag = createFragment();
		const wrap = frag.createDiv("pcal-newmail-notice");
		wrap.createDiv({ cls: "pcal-newmail-from", text: title });
		wrap.createDiv({ cls: "pcal-newmail-subject", text: body });
		const notice = new Notice(frag, 8000);
		// clicking the notice opens the newest of them, which is what the
		// hand is already reaching for
		notice.containerEl.addClass("pcal-newmail-clickable");
		notice.containerEl.addEventListener("click", () => {
			notice.hide();
			void this.openMailView().then(() => {
				const v = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
				if (v instanceof MailView) v.revealMessage(arrivals[arrivals.length - 1]);
			});
		});

		// a desktop notification as well when the window is not in front,
		// since a Notice inside a hidden window helps nobody
		if (Platform.isDesktopApp && !document.hasFocus()) {
			try {
				new Notification(title, { body, silent: false });
			} catch {
				/* the OS refused or the permission is denied: the Notice stands */
			}
		}
	}

	/** The unread count on the mail ribbon icon. */
	paintRibbonBadge() {
		const el = this.ribbonEl;
		if (!el) return;
		const n = this.mailAccounts().reduce((sum, a) => sum + (this.mailCache.get(a.id)?.messages.filter((m) => m.unread).length ?? 0), 0);
		el.toggleClass("pcal-has-unread", n > 0 && this.settings.mailBadge);
		el.setAttribute("data-pcal-unread", n > 99 ? "99+" : String(n));
	}

	/* ----- focus mode ----- */

	/** Whether the vault's page tree is folded away to give mail the width.
	 *
	 *  Only the left sidebar moves, and only through Obsidian's own API, so
	 *  the app knows it is closed and its own toggles keep working. The
	 *  ribbon stays put: it is a handful of pixels, and while it is there the
	 *  window can never feel like a mode you cannot leave. Clicking anything
	 *  on it brings the tree straight back, which is what it is for. */
	private pagesHidden = false;
	/** Whether the tree is folded away right now, which is not the same
	 *  question. The fold belongs to the mail and calendar tabs: step off one
	 *  and the tree comes back, step onto one and it folds again. Without the
	 *  distinction, turning it on and then opening a note left the pages
	 *  collapsed behind a note that had every use for them and no button to
	 *  bring them back — a mode you could not leave, which is exactly what the
	 *  ribbon escape hatch existed to avoid and did not really solve. */
	private pagesFolded = false;

	focusOn(): boolean {
		return this.pagesHidden;
	}

	/** Whether the tab in front is one the fold is for. */
	private onPowerView(): boolean {
		const v = this.app.workspace.getActiveViewOfType(ItemView);
		return v instanceof MailView || v instanceof PowerCalendarView;
	}

	/** Fold the whole left side away, and put it back.
	 *
	 *  Nothing but collapse and expand: an earlier version also stored the
	 *  width and reapplied it, which squashed a left sidebar holding two
	 *  panes side by side into one, because a fixed width on the container
	 *  is a fixed width for both of them together. Obsidian remembers its own
	 *  layout perfectly well; the job here is only to shut it and open it. */
	/** Told when the fold changes, so the copy of the button in the tab you are
	 *  not on stops claiming otherwise. Deliberately not the redraw listeners:
	 *  nothing drawn depends on this, and repainting every row is what made
	 *  the toggle feel like it had hung. */
	private focusWatchers = new Set<() => void>();

	watchFocus(fn: () => void): () => void {
		this.focusWatchers.add(fn);
		return () => this.focusWatchers.delete(fn);
	}

	toggleFocus(on?: boolean) {
		this.pagesHidden = on ?? !this.pagesHidden;
		this.applyFocus();
		for (const fn of this.focusWatchers) fn();
		// deliberately no notify(): nothing drawn depends on this, and a
		// repaint of every row is what made the toggle feel like it hung.
		// The button that was clicked updates its own state.
	}

	/** Put the tree where the current tab wants it: away while the fold is on
	 *  and mail or the calendar is in front, back for anything else.
	 *
	 *  Only ever moves it when the answer has actually changed, so opening a
	 *  note with the fold off never touches a sidebar the user arranged, and
	 *  the leaf changes that arrive several at a time cost one call each. */
	applyFocus() {
		const want = this.pagesHidden && this.onPowerView();
		if (want === this.pagesFolded) return;
		const left = this.app.workspace.leftSplit;
		if (!left) return;
		if (want) left.collapse();
		else left.expand();
		this.pagesFolded = want;
	}

	/** Hand the tree back this instant, without disarming the fold.
	 *
	 *  The ribbon click that calls this is on its way to opening something in
	 *  the tree and must not fire into a pane that is still shut, and it runs
	 *  before the leaf change that would otherwise sort this out. */
	private unfoldNow() {
		if (!this.pagesFolded) return;
		this.app.workspace.leftSplit?.expand();
		this.pagesFolded = false;
	}

	/* ----- automatic replies ----- */

	private autoReplyCache = new Map<string, AutoReply>();

	autoReplyFor(accountId: string): AutoReply | null {
		return this.autoReplyCache.get(accountId) ?? null;
	}

	/** Accounts currently answering mail on your behalf, for the banner that
	 *  stops an out-of-office running for three weeks after you got back. */
	autoReplyingAccounts(): GraphAccount[] {
		return this.mailAccounts().filter((a) => {
			const s = this.autoReplyCache.get(a.id);
			if (!s || s.status === "disabled") return false;
			if (s.status === "alwaysEnabled") return true;
			const start = fromGraphDateTime(s.scheduledStartDateTime);
			const end = fromGraphDateTime(s.scheduledEndDateTime);
			const now = Date.now();
			return (start == null || start <= now) && (end == null || end > now);
		});
	}

	async loadAutoReply(accountId: string): Promise<AutoReply | null> {
		const a = this.accountById(accountId);
		if (!a || !a.grantedScope.includes("MailboxSettings")) return null;
		try {
			const s = await getAutoReply(await this.graphTokenFor(a));
			if (s) this.autoReplyCache.set(accountId, s);
			this.notify();
			return s;
		} catch {
			return null;
		}
	}

	async saveAutoReply(accountId: string, setting: AutoReply): Promise<boolean> {
		const a = this.accountById(accountId);
		if (!a) return false;
		try {
			await setAutoReply(await this.graphTokenFor(a), setting);
			this.autoReplyCache.set(accountId, setting);
			this.notify();
			new Notice(setting.status === "disabled" ? "Power Desk: automatic replies are off." : "Power Desk: automatic replies are on.");
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** Mark everything unread in a folder as read.
	 *
	 *  Graph has no bulk read flag, so this is one write per message, twenty
	 *  at a time. It reports what it did rather than claiming the folder is
	 *  clear, because a folder that gained mail while this ran is not. */
	async markFolderRead(accountId: string, folderId: string, onProgress: (done: number, total: number) => void): Promise<{ done: number; failed: number; complete: boolean }> {
		const a = this.accountById(accountId);
		if (!a) return { done: 0, failed: 0, complete: true };
		try {
			const token = await this.graphTokenFor(a);
			const found = await listUnreadIdsInFolder(token, folderId);
			let done = 0;
			let failed = 0;
			for (const group of chunk(found.ids, GRAPH_BATCH_MAX)) {
				const r = await patchReadBatch(token, group, true);
				done += r.ok.length;
				failed += r.failed.length;
				onProgress(done + failed, found.ids.length);
			}
			// the cached lists still say unread until they are told otherwise
			const marked = new Set(found.ids);
			for (const list of this.cachedMailLists(accountId)) for (const m of list.messages) if (marked.has(m.id)) m.unread = false;
			const st = this.mailFolderCache.get(accountId);
			if (st) st.fetchedAt = 0;
			this.ensureMailFolders(true);
			this.notify();
			return { done, failed, complete: found.complete };
		} catch (e) {
			this.graphErrorNotice(e);
			return { done: 0, failed: 0, complete: true };
		}
	}

	/* ----- journal ----- */

	/** Sent Items, fetched into the folder cache so the journal can say what
	 *  went out as well as what came in. */
	async ensureSentForJournal(): Promise<void> {
		for (const a of this.mailAccounts()) this.ensureFolderMail(a.id, "sentitems", false);
	}

	sentMail(): PCMail[] {
		const out: PCMail[] = [];
		for (const a of this.mailAccounts()) out.push(...this.folderMail(a.id, "sentitems"));
		return out;
	}

	/** Put a day's journal into that day's daily note.
	 *
	 *  Appended rather than written over, and appended only once: running it
	 *  twice should not give you the morning twice. An existing note keeps
	 *  everything already in it, since the daily note is the user's writing
	 *  and this is a guest in it. */
	async appendToDailyNote(key: string, markdown: string): Promise<boolean> {
		const folder = normalizePath(this.settings.notesFolder.trim() || "Calendar");
		await this.ensureFolder(folder);
		const path = normalizePath(`${folder}/${key}.md`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		try {
			if (existing instanceof TFile) {
				const head = markdown.split("\n")[0];
				let already = false;
				// read and write in one atomic step, so an edit landing between
				// the duplicate check and the append cannot be lost
				await this.app.vault.process(existing, (data) => {
					already = data.includes(head);
					return already ? data : `${data.trimEnd()}\n\n${markdown}`;
				});
				if (already) new Notice("Power Desk: this day is already in that note.");
				await this.showNote(existing);
				return true;
			}
			const f = await this.app.vault.create(path, markdown);
			await this.showNote(f);
			return true;
		} catch (e) {
			new Notice("Power Desk: could not write that note. " + (e instanceof Error ? e.message : String(e)));
			return false;
		}
	}

	/* ----- Outlook notes ----- */

	private notesFolderId: string | null = null;
	private notesList: StickyNote[] = [];

	stickyNotes(): StickyNote[] {
		return this.notesList;
	}

	private notesAccount(): GraphAccount | null {
		return this.mailAccounts()[0] ?? null;
	}

	async loadStickyNotes(): Promise<{ ok: boolean; reason: string }> {
		const a = this.notesAccount();
		if (!a) return { ok: false, reason: "Connect a Microsoft account to see your Outlook notes." };
		try {
			const token = await this.graphTokenFor(a);
			this.notesFolderId = this.notesFolderId ?? (await findNotesFolder(token));
			if (!this.notesFolderId) return { ok: false, reason: "This mailbox has no Notes folder, so there are no Outlook notes to show." };
			this.notesList = await listStickyNotes(token, this.notesFolderId);
			this.notify();
			return { ok: true, reason: "" };
		} catch (e) {
			return { ok: false, reason: e instanceof Error ? e.message : String(e) };
		}
	}

	async stickyNoteBody(id: string): Promise<string> {
		const a = this.notesAccount();
		if (!a) return "";
		try {
			return await getStickyNoteBody(await this.graphTokenFor(a), id);
		} catch (e) {
			this.graphErrorNotice(e);
			return "";
		}
	}

	async addStickyNote(text: string): Promise<boolean> {
		const a = this.notesAccount();
		if (!a || !this.notesFolderId || !text.trim()) return false;
		try {
			await createStickyNote(await this.graphTokenFor(a), this.notesFolderId, text.trim());
			await this.loadStickyNotes();
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	/** A note is an item in a mail folder, so deleting one is deleting a
	 *  message: it lands in Deleted Items and can be recovered. */
	async removeStickyNote(id: string): Promise<void> {
		const a = this.notesAccount();
		if (!a) return;
		try {
			await deleteMessage(await this.graphTokenFor(a), id);
			this.notesList = this.notesList.filter((n) => n.id !== id);
			this.notify();
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/** An Outlook note into the vault, which is the thing this plugin can do
	 *  that Outlook cannot. */
	async stickyNoteToVault(n: StickyNote): Promise<void> {
		const body = (await this.stickyNoteBody(n.id)) || n.preview;
		const folder = normalizePath(this.settings.mailNotesFolder.trim() || this.settings.notesFolder.trim() || "Calendar");
		await this.ensureFolder(folder);
		const path = normalizePath(`${folder}/${sanitizeName(`${keyOfMs(n.changedMs || Date.now())} ${n.title}`)}.md`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.showNote(existing);
			return;
		}
		const f = await this.app.vault.create(path, `# ${n.title}\n\n${stripHtml(body)}\n`);
		await this.showNote(f);
	}

	/* ----- tasks ----- */

	private todoLists: TodoList[] = [];
	private todoTasks = new Map<string, TodoTask[]>();
	private todoAccount = "";

	todoListsFor(): TodoList[] {
		return this.todoLists;
	}

	tasksIn(listId: string): TodoTask[] {
		return this.todoTasks.get(listId) ?? [];
	}

	tasksNeedReconnect(): boolean {
		const accounts = this.mailAccounts();
		return accounts.length > 0 && !accounts.some((a) => a.grantedScope.includes("Tasks"));
	}

	/** The account that owns the task lists: the first that granted the
	 *  permission. To Do is per mailbox, and juggling several would be a
	 *  different feature than the one anybody asked for. */
	private taskAccount(): GraphAccount | null {
		return this.mailAccounts().find((a) => a.grantedScope.includes("Tasks")) ?? null;
	}

	async loadTaskLists(): Promise<TodoList[]> {
		const a = this.taskAccount();
		if (!a) return [];
		try {
			this.todoLists = await listTodoLists(await this.graphTokenFor(a));
			this.todoAccount = a.id;
			this.notify();
		} catch (e) {
			this.graphErrorNotice(e);
		}
		return this.todoLists;
	}

	async loadTasks(listId: string): Promise<void> {
		const a = this.taskAccount();
		if (!a) return;
		try {
			this.todoTasks.set(listId, await listTodoTasks(await this.graphTokenFor(a), listId));
			this.notify();
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async addTask(listId: string, title: string, dueMs: number | null): Promise<void> {
		const a = this.taskAccount();
		if (!a || !title.trim()) return;
		try {
			await createTodoTask(await this.graphTokenFor(a), listId, title.trim(), dueMs);
			await this.loadTasks(listId);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async setTaskDone(t: TodoTask, done: boolean): Promise<void> {
		const a = this.taskAccount();
		if (!a) return;
		// the tick lands before the mailbox answers, like every other action
		const list = this.todoTasks.get(t.listId);
		const hit = list?.find((x) => x.id === t.id);
		if (hit) hit.done = done;
		this.notify();
		try {
			await setTodoTaskDone(await this.graphTokenFor(a), t.listId, t.id, done);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async removeTask(t: TodoTask): Promise<void> {
		const a = this.taskAccount();
		if (!a) return;
		try {
			await deleteTodoTask(await this.graphTokenFor(a), t.listId, t.id);
			await this.loadTasks(t.listId);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/** Flagged mail, which is what Outlook's own task list shows beside the
	 *  real tasks: a message you flagged is a thing to do. */
	flaggedMail(): PCMail[] {
		const out: PCMail[] = [];
		for (const a of this.mailAccounts()) for (const list of this.cachedMailLists(a.id)) for (const m of list.messages) if (m.flagged && !out.some((x) => x.id === m.id)) out.push(m);
		return out.sort((x, y) => y.receivedMs - x.receivedMs);
	}

	/* ----- categories ----- */

	private categoryCache = new Map<string, OutlookCategory[]>();

	categoriesFor(accountId: string): OutlookCategory[] {
		return this.categoryCache.get(accountId) ?? [];
	}

	/** The mailbox's category list, read once per session per account. */
	async loadCategories(accountId: string): Promise<OutlookCategory[]> {
		if (this.categoryCache.has(accountId)) return this.categoryCache.get(accountId) ?? [];
		const a = this.accountById(accountId);
		if (!a || !a.grantedScope.includes("MailboxSettings")) return [];
		try {
			const cats = await listMasterCategories(await this.graphTokenFor(a));
			this.categoryCache.set(accountId, cats);
			this.notify();
			return cats;
		} catch {
			// no categories is a usable state; the menu simply says so
			this.categoryCache.set(accountId, []);
			return [];
		}
	}

	/** The first of Outlook's colors this mailbox is not already using, so a new
	 *  category is distinguishable without asking for a second decision. */
	freeCategoryColor(accountId: string): string {
		const used = new Set(this.categoriesFor(accountId).map((x) => x.color));
		for (let i = 0; i < 25; i++) if (!used.has(`preset${i}`)) return `preset${i}`;
		return "preset0";
	}

	/** Whether this mailbox will answer questions about its categories at all.
	 *  The list, and making one, both read under MailboxSettings, which is the
	 *  permission inbox rules already ask for. */
	canReadCategories(accountId: string): boolean {
		return !!this.accountById(accountId)?.grantedScope.includes("MailboxSettings");
	}

	/** The color a category should be drawn in, from its mailbox's list. */
	categoryColorFor(accountId: string, name: string): string {
		const hit = this.categoriesFor(accountId).find((c) => c.displayName.toLowerCase() === name.toLowerCase());
		return categoryColor(hit?.color ?? "");
	}

	/** Make a category in the mailbox's list, so it can be used everywhere. */
	async newCategory(accountId: string, name: string, color: string): Promise<boolean> {
		const a = this.accountById(accountId);
		if (!a || !name.trim()) return false;
		try {
			await createCategory(await this.graphTokenFor(a), name.trim(), color);
			this.categoryCache.delete(accountId);
			await this.loadCategories(accountId);
			return true;
		} catch (e) {
			this.graphErrorNotice(e);
			return false;
		}
	}

	async recolorCategory(accountId: string, categoryId: string, color: string): Promise<void> {
		const a = this.accountById(accountId);
		if (!a) return;
		try {
			await updateCategoryColor(await this.graphTokenFor(a), categoryId, color);
			this.categoryCache.delete(accountId);
			await this.loadCategories(accountId);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async removeCategory(accountId: string, categoryId: string): Promise<void> {
		const a = this.accountById(accountId);
		if (!a) return;
		const gone = this.categoriesFor(accountId).find((c) => c.id === categoryId)?.displayName;
		try {
			await deleteCategory(await this.graphTokenFor(a), categoryId);
			// a category pinned to Favorites goes with it: the mail keeps the
			// label, but there is no longer a category to open, and a pin that
			// opened an ever-emptying list would be worse than none
			if (gone) this.repointCategoryFavorite(accountId, gone, null);
			this.categoryCache.delete(accountId);
			await this.loadCategories(accountId);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/** Follow a category's pin through a rename, or drop it on a delete.
	 *  Favorites store a folder id, and a category's id is its name, so a name
	 *  that changes leaves the pin pointing at nothing without this. */
	private repointCategoryFavorite(accountId: string, from: string, to: string | null) {
		const oldId = categoryFolderId(from);
		const s = this.settings;
		if (!s.mailFavorites.some((f) => f.accountId === accountId && f.folderId === oldId)) return;
		s.mailFavorites = to
			? s.mailFavorites.map((f) => (f.accountId === accountId && f.folderId === oldId ? { ...f, folderId: categoryFolderId(to) } : f))
			: s.mailFavorites.filter((f) => !(f.accountId === accountId && f.folderId === oldId));
		this.queueSave();
		this.notify();
	}

	/** How much mail a category is on, before anything is changed. */
	async countCategoryUse(accountId: string, name: string): Promise<{ hits: { id: string; categories: string[] }[]; complete: boolean } | null> {
		const a = this.accountById(accountId);
		if (!a) return null;
		try {
			return await findMessagesByCategory(await this.graphTokenFor(a), name);
		} catch (e) {
			this.graphErrorNotice(e);
			return null;
		}
	}

	/** Retag every message from one category to another, then retire the old
	 *  one. The nearest thing to a rename the API allows.
	 *
	 *  The order matters and is the whole safety story: the new category is
	 *  made first, every message is rewritten, and the old category is
	 *  deleted only if every last rewrite succeeded. A partial run therefore
	 *  leaves the old category in place and both names on the list, which is
	 *  untidy but recoverable by running it again. Deleting first, or
	 *  deleting after a partial run, would strip the color from whatever was
	 *  left behind and give no way to find it. */
	async replaceCategoryEverywhere(
		accountId: string,
		from: { id: string; displayName: string; color: string },
		to: string,
		hits: { id: string; categories: string[] }[],
		onProgress: (done: number, total: number) => void
	): Promise<{ changed: number; failed: number; retired: boolean }> {
		const a = this.accountById(accountId);
		if (!a) return { changed: 0, failed: hits.length, retired: false };
		const token = await this.graphTokenFor(a);
		// the new name inherits the old one's color, so nothing looks different
		// afterwards except the word
		if (!this.categoriesFor(accountId).some((c) => c.displayName.toLowerCase() === to.toLowerCase())) {
			try {
				await createCategory(token, to, from.color);
			} catch (e) {
				this.graphErrorNotice(e);
				return { changed: 0, failed: hits.length, retired: false };
			}
		}
		let changed = 0;
		let failed = 0;
		for (const group of chunk(hits, GRAPH_BATCH_MAX)) {
			const items = group.map((h) => ({ id: h.id, categories: replaceCategory(h.categories, from.displayName, to) }));
			const r = await patchCategoriesBatch(token, items);
			changed += r.ok.length;
			failed += r.failed.length;
			onProgress(changed + failed, hits.length);
		}
		let retired = false;
		if (!failed) {
			try {
				await deleteCategory(token, from.id);
				retired = true;
			} catch (e) {
				this.graphErrorNotice(e);
			}
		}
		// a pinned category follows the rename rather than being left pointing
		// at a name the mailbox no longer has. It moves even when the old one
		// could not be retired, since the mail is under the new name either way
		this.repointCategoryFavorite(accountId, from.displayName, to);
		this.categoryCache.delete(accountId);
		await this.loadCategories(accountId);
		// the lists hold the old names until they are read again
		this.refreshMailAll(true);
		return { changed, failed, retired };
	}

	/** Write a message's categories, list and all, and show it at once. */
	async setMailCategories(m: PCMail, categories: string[]): Promise<void> {
		const a = this.accountById(m.accountId);
		if (!a) return;
		for (const list of this.cachedMailLists(a.id)) {
			const cached = list.messages.find((x) => x.id === m.id);
			if (cached) cached.categories = categories.length ? categories : undefined;
		}
		this.notify();
		try {
			await setMessageCategories(await this.graphTokenFor(a), m.id, categories);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/* ----- sender photos ----- */

	/** Address to photo, with null meaning "asked, they have none", which is
	 *  what stops a mailbox full of newsletters asking forever. */
	private photoCache = new Map<string, string | null>();
	private photoInFlight = new Set<string>();
	/** Set when the mailbox says the photo permission was never granted, so
	 *  the whole idea is dropped rather than retried per sender. */
	private photosDenied = false;

	photoFor(address: string): string | null {
		return this.photoCache.get((address ?? "").toLowerCase()) ?? null;
	}

	/** Fetch the photos of everyone in a rendered list who is still unknown.
	 *  Capped per pass: scrolling a long list should not open two hundred
	 *  requests, and the ones that matter are the ones on screen. */
	ensurePhotos(addresses: string[]): void {
		if (this.photosDenied || !this.settings.mailPhotos) return;
		const a = this.mailAccounts().find((x) => x.grantedScope.includes("ProfilePhoto"));
		if (!a) return;
		let budget = 12;
		for (const raw of addresses) {
			const addr = (raw ?? "").toLowerCase().trim();
			if (!addr.includes("@") || this.photoCache.has(addr) || this.photoInFlight.has(addr)) continue;
			if (budget-- <= 0) break;
			this.photoInFlight.add(addr);
			void (async () => {
				try {
					const url = await getUserPhoto(await this.graphTokenFor(a), addr);
					this.photoCache.set(addr, url);
					if (url) {
						this.notify();
						this.queueCachePersist();
					}
				} catch (e) {
					// the permission is missing: stop, rather than failing
					// once per sender for the rest of the session
					if (e instanceof Error && e.message === "no-photo-permission") this.photosDenied = true;
					else this.photoCache.set(addr, null);
				} finally {
					this.photoInFlight.delete(addr);
				}
			})();
		}
	}

	/* ----- folder management ----- */

	/** Make a folder, at the root or inside another, and show it straight
	 *  away by refetching the tree. */
	async newFolder(accountId: string, name: string, parentId: string | null): Promise<string | null> {
		const a = this.accountById(accountId);
		if (!a || !name.trim()) return null;
		try {
			const id = await createMailFolder(await this.graphTokenFor(a), name.trim(), parentId);
			await this.refetchFolders(a);
			new Notice(`Power Desk: created ${name.trim()}.`);
			return id;
		} catch (e) {
			this.graphErrorNotice(e);
			return null;
		}
	}

	async renameFolder(accountId: string, folderId: string, name: string): Promise<void> {
		const a = this.accountById(accountId);
		if (!a || !name.trim()) return;
		try {
			await renameMailFolder(await this.graphTokenFor(a), folderId, name.trim());
			await this.refetchFolders(a);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	async removeFolder(accountId: string, folderId: string): Promise<void> {
		const a = this.accountById(accountId);
		if (!a) return;
		try {
			await deleteMailFolder(await this.graphTokenFor(a), folderId);
			// forget everything that pointed at it, so nothing keeps a dead id
			this.folderCache.delete(`${accountId}:${folderId}`);
			this.settings.mailFavorites = this.settings.mailFavorites.filter((f) => !(f.accountId === accountId && f.folderId === folderId));
			this.settings.mailRecentFolders = this.settings.mailRecentFolders.filter((f) => !(f.accountId === accountId && f.folderId === folderId));
			this.settings.mailHiddenFolders = this.settings.mailHiddenFolders.filter((f) => !(f.accountId === accountId && f.folderId === folderId));
			this.settings.mailFolderOrder = this.settings.mailFolderOrder.map((o) => (o.accountId === accountId ? { ...o, folderIds: o.folderIds.filter((id) => id !== folderId) } : o));
			this.queueSave();
			await this.refetchFolders(a);
		} catch (e) {
			this.graphErrorNotice(e);
		}
	}

	/** Pull the tree again after a change, bypassing the staleness rules so
	 *  the pane shows what just happened rather than what it remembers. */
	private async refetchFolders(a: GraphAccount): Promise<void> {
		const st = this.mailFolderCache.get(a.id);
		if (st) st.fetchedAt = 0;
		this.ensureMailFolders(true);
		this.notify();
	}

	/** A folder and everything under it, for a delete confirmation that can
	 *  say how much is actually at stake. */
	folderWeight(accountId: string, folderId: string): { messages: number; folders: number } {
		const st = this.mailFolderCache.get(accountId);
		if (!st) return { messages: 0, folders: 0 };
		const ids = folderSubtreeIds(st.folders, folderId);
		let messages = 0;
		for (const f of st.folders) if (ids.has(f.id)) messages += f.total;
		return { messages, folders: Math.max(0, ids.size - 1) };
	}

	/* ----- filing ----- */

	/** File messages into a folder. Graph moves within one mailbox, so this
	 *  takes the messages of a single account; the picker is what decides
	 *  which those are.
	 *
	 *  The rows leave the list before the mailbox has answered, because
	 *  filing is a rhythm and waiting on a round trip per message would break
	 *  it. A move that fails puts its message back on the next refresh, which
	 *  is the same recovery every other optimistic action here relies on. */
	async moveMail(targets: PCMail[], accountId: string, folderId: string, folderName: string): Promise<number> {
		const a = this.accountById(accountId);
		const mine = targets.filter((t) => t.accountId === accountId);
		if (!a || !mine.length) return 0;
		const ids = new Set(mine.map((m) => m.id));
		for (const list of this.cachedMailLists(accountId)) list.messages = list.messages.filter((x) => !ids.has(x.id));
		this.rememberFolder(accountId, folderId, folderName);
		// whatever is showing the destination is now wrong, so make it refetch
		const dest = this.folderCache.get(`${accountId}:${folderId}`);
		if (dest) dest.fetchedAt = 0;
		this.notify();
		let moved = 0;
		try {
			const token = await this.graphTokenFor(a);
			for (const m of mine) {
				await moveMessage(token, m.id, folderId);
				moved++;
			}
		} catch (e) {
			this.graphErrorNotice(e);
		}
		this.queueSave();
		return moved;
	}

	/** Keep the last few folders filed into, newest first. */
	private rememberFolder(accountId: string, folderId: string, name: string) {
		const next = [{ accountId, folderId, name }, ...this.settings.mailRecentFolders.filter((r) => !(r.accountId === accountId && r.folderId === folderId))];
		this.settings.mailRecentFolders = next.slice(0, 8);
	}

	/** Every folder that can be filed into, for the move picker: the accounts
	 *  given, their trees flattened, with the path so two folders called
	 *  "Archive" under different parents can be told apart. */
	moveTargets(accountIds: string[]): { accountId: string; accountLabel: string; folderId: string; name: string; path: string }[] {
		const out: { accountId: string; accountLabel: string; folderId: string; name: string; path: string }[] = [];
		for (const a of this.mailAccounts()) {
			if (accountIds.length && !accountIds.includes(a.id)) continue;
			const label = this.nameOf(a);
			const byId = new Map(this.folderTreeFor(a).map((n) => [n.folder.id, n.folder]));
			for (const { folder } of this.folderTreeFor(a)) {
				// the path walks up the parents, so nested folders read fully
				const parts: string[] = [folder.name];
				let p = folder.parentId;
				for (let i = 0; p && i < 12; i++) {
					const up = byId.get(p);
					if (!up) break;
					parts.unshift(up.name);
					p = up.parentId;
				}
				out.push({ accountId: a.id, accountLabel: label, folderId: folder.id, name: folder.name, path: parts.join(" / ") });
			}
		}
		return out;
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

	/** The signature to put on a message, as HTML.
	 *
	 *  `accountId` is which mailbox is sending and `kind` is whether this is
	 *  a first message or a reply, because the long block belongs on one and
	 *  rarely on the fourth reply of a thread. With no account named, the
	 *  first sending account's setting stands in, which is what a new message
	 *  written before a sender is chosen needs. */
	signatureHtml(accountId?: string, kind: "new" | "reply" = "new"): string {
		const id = accountId ?? this.mailAccounts()[0]?.id ?? "";
		const sig = signatureFor(this.settings.mailSignatures, this.settings.mailSignatureUse, id, kind);
		const html = (sig?.html ?? "").trim();
		if (!html) return "";
		if (/<[a-z][\s\S]*>/i.test(html)) return html;
		const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
		return `<p>${esc(html).replace(/\n/g, "<br>")}</p>`;
	}

	/** Put a newly shipped action on an existing toolbar, once ever. */
	private offerToolbarActions() {
		const offers: { key: string; get: () => string[]; set: (v: string[]) => void; id: string }[] = [
			{
				key: "calendar:print",
				id: "print",
				get: () => this.settings.calendarToolbar,
				set: (v) => (this.settings.calendarToolbar = v),
			},
			{
				key: "calendar:searchEvents",
				id: "searchEvents",
				get: () => this.settings.calendarToolbar,
				set: (v) => (this.settings.calendarToolbar = v),
			},
			// one Print on the mail toolbar, not two: the window it opens
			// carries both styles, so a second button would only be a longer
			// way to the same place
			{
				key: "mail:print",
				id: "print",
				get: () => this.settings.mailToolbar,
				set: (v) => (this.settings.mailToolbar = v),
			},
		];
		let changed = false;
		for (const o of offers) {
			if (this.settings.toolbarAdded.includes(o.key)) continue;
			this.settings.toolbarAdded = [...this.settings.toolbarAdded, o.key];
			changed = true;
			if (!o.get().includes(o.id)) o.set([...o.get(), o.id]);
		}
		if (changed) this.queueSave();
	}

	/** Carry the old single signature into the named list, once. */
	private migrateSignatures() {
		const moved = migrateSignature(this.settings.mailSignature, this.settings.mailSignatures, this.settings.graphAccounts.map((a) => a.id), freshId());
		if (!moved) return;
		this.settings.mailSignatures = moved.sigs;
		this.settings.mailSignatureUse = moved.use;
		this.queueSave();
	}

	/** Which signature an account uses, creating the row on first ask. */
	signatureUseFor(accountId: string): SignatureUse {
		const hit = this.settings.mailSignatureUse.find((x) => x.accountId === accountId);
		if (hit) return hit;
		const fresh = { accountId, newId: "", replyId: "" };
		this.settings.mailSignatureUse = [...this.settings.mailSignatureUse, fresh];
		return fresh;
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

	/** Send a composed message through a draft.
	 *
	 *  Attachments and a send time are both properties of a message, so
	 *  either one means the message has to exist before it can go: create a
	 *  draft, hang the files and the time on it, then send. A reply already
	 *  has its draft, so it only gets updated. Plain new mail with neither
	 *  skips all of this and posts straight to sendMail, which is one request
	 *  instead of three and keeps the Power Assistant transport working. */
	async sendComposed(o: {
		accountId: string | null;
		draftId: string | null;
		preferAccountId?: string;
		patch: { subject: string; html: string; to: string[]; cc: string[]; bcc?: string[]; receipts?: Receipts };
		files: OutgoingFile[];
		whenMs?: number;
	}): Promise<boolean> {
		const preferred = o.accountId ? this.accountById(o.accountId) : o.preferAccountId ? this.accountById(o.preferAccountId) : null;
		const a =
			(preferred && preferred.refresh && preferred.grantedScope.includes("Mail.Send") ? preferred : null) ??
			this.settings.graphAccounts.find((x) => !!x.refresh && x.grantedScope.includes("Mail.Send")) ??
			null;
		if (!a) {
			new Notice("Power Desk: this needs a Microsoft account connected with send permission.");
			return false;
		}
		try {
			const token = await this.graphTokenFor(a);
			// a signature's images live as data urls while being edited; mail
			// clients refuse those, so they go as inline attachments instead
			const { html, images } = extractInlineImages(o.patch.html, freshId());
			const patch = { ...o.patch, html };
			let draftId = o.draftId;
			if (draftId) await updateDraft(token, draftId, patch);
			else draftId = await createDraftMessage(token, { to: patch.to, cc: patch.cc, bcc: patch.bcc, subject: patch.subject, html: patch.html, receipts: patch.receipts });
			if (!draftId) return false;
			for (const img of images) await addInlineImage(token, draftId, img);
			for (const f of o.files) await addFileAttachment(token, draftId, f);
			if (o.whenMs) await setDeferredSend(token, draftId, o.whenMs);
			await sendDraft(token, draftId);
			new Notice(o.whenMs ? `Power Desk: scheduled for ${fmtWhen(o.whenMs, this.settings.use24h)}.` : "Power Desk: mail sent.");
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
	mailSender(preferAccountId?: string): { label: string; send: (m: { to: string[]; cc?: string[]; bcc?: string[]; subject: string; html: string; receipts?: Receipts }) => Promise<void> } | null {
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

	/** A message into the vault: dated note, the message as Markdown with its
	 *  pictures, source link. */
	async saveMailToNote(m: PCMail): Promise<void> {
		const body = await this.readMailBody(m);
		// mail has its own folder so it can be filed apart from event notes (and
		// into a protected folder); empty falls back to the calendar folder
		const folder = normalizePath(this.settings.mailNotesFolder.trim() || this.settings.notesFolder.trim() || "Calendar");
		await this.ensureFolder(folder);
		const base = sanitizeName(`${keyOfMs(m.receivedMs)} ${m.subject}`);
		const path = normalizePath(`${folder}/${base}.md`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.showNote(existing);
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
			await this.mailNoteBody(m, body, folder, base, path),
			"",
		];
		const f = await this.app.vault.create(path, lines.join("\n"));
		await this.showNote(f);
		new Notice("Power Desk: mail saved to a note.");
	}

	/** A saved message's body. HTML mail becomes Markdown, so the headings,
	 *  lists, links and pictures that are the message survive the trip into the
	 *  vault; the stripped text stays the fallback for plain-text mail and for
	 *  anything the conversion makes nothing of. */
	private async mailNoteBody(m: PCMail, body: { text: string; html?: string } | null, folder: string, base: string, notePath: string): Promise<string> {
		const html = body?.html?.trim() ?? "";
		const fallback = body?.text?.trim() || m.preview;
		if (!html) return fallback;
		const embeds = this.settings.mailNoteImages ? await this.saveInlineImages(m, html, folder, base, notePath) : new Map<string, MailEmbed>();
		return mailBodyMarkdown(html, htmlToMarkdown, embeds) || fallback;
	}

	/** A message's inline pictures written beside the note, and the embed for
	 *  each, keyed the way the Markdown pass looks them up.
	 *
	 *  They go into an attachments folder under the saved-mail folder rather
	 *  than the vault's own attachment folder so that a saved-mail folder under
	 *  Power Connect protection keeps the pictures encrypted too, and so a mail
	 *  note and its pictures move together.
	 *
	 *  Data URLs carry their own bytes. A cid: reference has to be fetched, and
	 *  Graph only reports contentId with the payload, so the inline attachments
	 *  are read and matched afterwards, the same walk the reading pane makes;
	 *  a sender who marked its pictures as ordinary attachments is picked up on
	 *  a second pass, which only runs when something is still missing. */
	private async saveInlineImages(m: PCMail, html: string, folder: string, base: string, notePath: string): Promise<Map<string, MailEmbed>> {
		const all = mailInlineImages(html);
		const refs = all.slice(0, MAX_NOTE_IMAGES);
		if (!refs.length) return new Map();
		// a cap that says nothing reads as "that was all of them"
		if (all.length > refs.length) new Notice(`Power Desk: keeping the first ${refs.length} pictures of ${all.length}.`);
		const bytes = new Map<string, { base64: string; ext: string }>();
		for (const r of refs) {
			const parsed = r.dataUrl ? parseDataUrl(r.dataUrl) : null;
			if (parsed) bytes.set(r.key, parsed);
		}
		const missing = () => refs.filter((r) => r.cid && !bytes.has(r.key));
		if (missing().length) {
			const atts = await this.mailAttachments(m);
			const take = async (list: MailAttachment[]) => {
				for (const raw of await Promise.all(list.slice(0, MAX_NOTE_IMAGES).map((a) => this.mailAttachmentRaw(m, a.id, true)))) {
					if (!raw?.contentId) continue;
					const key = normalizeCid(raw.contentId);
					if (refs.some((r) => r.key === key)) bytes.set(key, { base64: raw.contentBytes, ext: imageExtension(raw.name, raw.contentType) });
				}
			};
			await take(atts.filter((a) => a.isInline));
			if (missing().length) await take(atts.filter((a) => !a.isInline && /^image\//i.test(a.contentType)));
		}
		if (!bytes.size) return new Map();
		const dir = normalizePath(`${folder}/attachments`);
		await this.ensureFolder(dir);
		const out = new Map<string, MailEmbed>();
		let n = 0;
		for (const r of refs) {
			const hit = bytes.get(r.key);
			if (!hit) continue;
			try {
				const data = base64ToArrayBuffer(hit.base64);
				const f = await this.app.vault.createBinary(this.freeAttachmentPath(dir, `${base} ${++n}`, hit.ext), data);
				// generateMarkdownLink writes whatever this vault's link style is;
				// the "!" is what makes it an embed rather than a link to a file
				out.set(r.key, { link: `!${this.app.fileManager.generateMarkdownLink(f, notePath)}`, naturalWidth: imageSize(new Uint8Array(data))?.w });
			} catch {
				// one picture that will not write is not worth losing the note over
			}
		}
		return out;
	}

	/** A path in `dir` nothing occupies yet. */
	private freeAttachmentPath(dir: string, base: string, ext: string): string {
		for (let n = 0; n < 100; n++) {
			const p = normalizePath(`${dir}/${base}${n ? ` (${n})` : ""}.${ext}`);
			if (!this.app.vault.getAbstractFileByPath(p)) return p;
		}
		return normalizePath(`${dir}/${base} ${Date.now()}.${ext}`);
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
				const pc = graphMailToPC(o, a.id, this.nameOf(a), a.label);
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
					const body = await this.readMailBody(msg, true, "idle");
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

	/**
	 * The main-area tab already showing this path, if there is one.
	 *
	 * Asked through `getViewState()` rather than `leaf.view.file`, because every
	 * tab you are not standing in is deferred since 1.7.2: its view is a stand-in
	 * that holds no file, and reaching for one to ask would load every tab in the
	 * window. The view state carries the path whether the view is real or not.
	 *
	 * Main-area leaves only. A note showing in a sidebar is not a tab, and a note
	 * deliberately popped out into a window of its own should not have the
	 * calendar pulling focus to another window behind your back.
	 */
	private openLeafFor(path: string): WorkspaceLeaf | null {
		const hits: WorkspaceLeaf[] = [];
		this.app.workspace.iterateRootLeaves((leaf) => {
			const open = leaf.getViewState().state?.file;
			if (typeof open === "string" && open === path) hits.push(leaf);
		});
		return hits[0] ?? null;
	}

	/**
	 * Show a note: step to the tab already holding it, or open a fresh one.
	 *
	 * Opening the same day, mail, or event note twice used to hand you a second
	 * copy of it, because the open said only where to put a note and nothing
	 * about where that note already was: two scroll positions, two undo
	 * histories, and edits landing in whichever one you looked at last.
	 *
	 * The notesInNewTab setting still decides where a note lands the first time.
	 * It reads as "do not take over the tab I am in", and stepping to the tab the
	 * note is already in honors that just as well as making a new one, so a note
	 * that is already open is shown rather than opened again either way. The
	 * override is for the routes that never consulted the setting, because a base
	 * file and a shortcut are not the notes it is about.
	 */
	async showNote(f: TFile, newTab = this.settings.notesInNewTab): Promise<WorkspaceLeaf> {
		const open = this.openLeafFor(f.path);
		if (open) {
			await this.app.workspace.revealLeaf(open);
			this.app.workspace.setActiveLeaf(open, { focus: true });
			return open;
		}
		const leaf = this.app.workspace.getLeaf(newTab);
		await leaf.openFile(f);
		return leaf;
	}

	async openEventNote(ev: PCEvent) {
		if (ev.notePath) {
			const f = this.app.vault.getAbstractFileByPath(ev.notePath);
			if (f instanceof TFile) await this.showNote(f);
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
		if (file instanceof TFile) await this.showNote(file);
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
		if (f instanceof TFile) await this.showNote(f);
	}

	/** Drop a ready-made Bases table over the event-notes folder (Power Bases'
	 *  table when it is installed, the core one otherwise) and open it. */
	async createEventsBase() {
		const folder = normalizePath(this.settings.notesFolder.trim() || "Calendar");
		const path = normalizePath(`${folder}/Events.base`);
		const existing = this.app.vault.getAbstractFileByPath(path);
		if (existing instanceof TFile) {
			await this.showNote(existing, false);
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
		await this.showNote(f, false);
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
	private selectedBody: { text: string; html?: string; toLine: string; unsub?: UnsubscribeInfo } | null = null;
	private selectedAtts: MailAttachment[] | null = null;
	/** Inline images resolved from cid: references to data urls. */
	private inlineCids: Map<string, string> | null = null;
	/** The favorite row a drag started from, for drop reordering. */
	private favDragIdx: number | null = null;
	/** The messages currently being dragged onto a folder, if any. */
	private mailDrag: PCMail[] | null = null;
	/** The folder being dragged to a new position among its siblings. */
	private folderDrag: { accountId: string; folderId: string } | null = null;
	/** Ctrl and Shift clicks build a bulk selection for the toolbar actions. */
	private multiSel = new Set<string>();
	private selAnchorId: string | null = null;
	/** The list as last rendered, in order, for shift ranges and bulk lookup.
	 *  One entry per visible row: a collapsed conversation contributes its
	 *  newest message, an expanded one its head and then every child. */
	private lastList: PCMail[] = [];
	/** What each visible row acts on, parallel to lastList. A collapsed
	 *  conversation head answers for its whole thread, so ticking it or
	 *  shift-selecting across it takes the entire back-and-forth, exactly as
	 *  Outlook does. Every other row answers for itself alone. */
	private lastRows: PCMail[][] = [];
	/** Conversations the user has opened, by thread key. Kept on the view
	 *  rather than in settings: which threads are open is a reading position,
	 *  not a preference, and it should not follow you to another device. */
	private expanded = new Set<string>();
	/** The conversations behind the last render, for the reading pane's
	 *  thread strip. Empty whenever grouping is off. */
	private lastThreads: MailThread[] = [];
	/** Set when a folder is entered: the next render points the reading pane
	 *  at the list's first message. */
	private autoSelectPending = false;
	/** The reply box's state lives on the view, so a refresh or the body
	 *  arriving cannot wipe a half-typed draft. */
	private mailToolBtns: HTMLElement[] = [];
	private toolsEl!: HTMLElement;
	private oooEl!: HTMLElement;
	private readTimer: number | null = null;
	/** List-header state: one extra filter beside Unread, and the sort. */
	private extraFilter: "none" | "priority" | "flagged" | "tome" | "attachments" = "none";
	/** Select mode shows every checkbox unchecked, ready for picking. */
	private selectMode = false;
	private sortBy: "date" | "from" | "subject" = "date";
	private sortAsc = false;
	private renderQueued = false;
	/** What is drawn in the message column right now, keyed by message id, with
	 *  what each row was built from. A refresh whose answer is the same row
	 *  leaves that row alone, which at fifteen hundred messages is the
	 *  difference between remaking the list and touching nothing. */
	private drawnItems = new Map<string, { sig: string; el: Element }>();
	/** The same bargain for the two panes either side, which are cheap to
	 *  describe and expensive to draw: the folder tree, and the reading pane,
	 *  where a rebuild means re-sanitizing the whole message body. */
	private lastReadSig: string | null = null;
	private lastFolderSig: string | null = null;
	/** Debounce on the search box: local search is cheap, not free. */
	private searchTypeTimer: number | null = null;
	private readonly onData = () => this.queueRender();

	constructor(
		leaf: WorkspaceLeaf,
		private plugin: PowerDeskPlugin
	) {
		super(leaf);
		this.scope = new Scope(this.app.scope);
		// Every binding passes through the guard: while the caret is in the
		// search box or a compose field, a letter is a letter and has to
		// reach the field untouched. Returning true lets it through.
		const guard = (evt: KeyboardEvent, cb: () => void) => {
			const t = evt.target as HTMLElement | null;
			if (t && (t.instanceOf(HTMLInputElement) || t.instanceOf(HTMLTextAreaElement) || t.isContentEditable)) return true;
			cb();
			return false;
		};
		const key = (mods: Modifier[], k: string, cb: () => void) => (this.scope as Scope).register(mods, k, (evt) => guard(evt, cb));

		// moving through the list, both the Gmail letters and the arrows
		key([], "j", () => this.moveSelection(1));
		key([], "k", () => this.moveSelection(-1));
		key([], "ArrowDown", () => this.moveSelection(1));
		key([], "ArrowUp", () => this.moveSelection(-1));
		key([], "Enter", () => this.openSelected());
		key([], "u", () => this.goBackOrClear());
		key([], "Escape", () => this.escape());

		// acting on what is selected
		key([], "e", () => this.archiveSelection());
		key([], "r", () => this.composeFor("reply"));
		key([], "a", () => this.composeFor("replyAll"));
		key([], "f", () => this.composeFor("forward"));
		key([], "c", () => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open());
		key([], "x", () => this.toggleSelectionMark());
		key([], "s", () => this.flagSelection());
		key([], "b", () => this.snoozeSelection());
		key([], "g", () => this.categorizeMenu(this.multiTargets()));
		key([], "v", () => this.openMovePicker(this.multiTargets()));
		key(["Shift"], "u", () => this.markSelection(false));
		key(["Shift"], "i", () => this.markSelection(true));
		key([], "Delete", () => this.deleteSelection(false));
		key(["Shift"], "Delete", () => this.deleteSelection(true));

		// conversations open and close under the arrows, like a tree
		key([], "ArrowRight", () => this.setThreadOpen(true));
		key([], "ArrowLeft", () => this.setThreadOpen(false));
		key([], "o", () => this.setThreadOpen(!this.selectedThreadOpen()));

		// finding things
		key([], "/", () => this.focusSearch());
		key(["Mod"], "k", () => new MailPaletteModal(this.app, this.plugin, this).open());
		key(["Shift"], "/", () => new MailShortcutsModal(this.app).open());
	}

	/** Everything the palette can offer right now: every folder in every
	 *  mail account, then the actions that apply. Places come first because
	 *  jumping is the common case, and an empty query should read as a list
	 *  of where you can go. */
	paletteItems(): PaletteItem[] {
		const out: PaletteItem[] = [];
		const s = this.plugin.settings;
		const go = (accountId: string, folderId: string, name: string) => {
			this.plugin.clearMailSearch();
			this.folderSel = { accountId, folderId, name };
			this.screen = "list";
			this.autoSelectPending = true;
			this.render();
		};
		out.push({
			label: "All inboxes",
			hint: "every account",
			terms: "unified inbox home",
			run: () => {
				this.plugin.clearMailSearch();
				this.folderSel = null;
				this.screen = "list";
				this.autoSelectPending = true;
				this.render();
			},
		});
		for (const a of this.plugin.mailAccounts()) {
			const label = this.plugin.nameOf(a);
			out.push({ label: "Unread Mail", hint: label, terms: "unread search folder", run: () => go(a.id, UNREAD_FOLDER, "Unread Mail") });
			for (const { folder } of this.plugin.folderTreeFor(a)) {
				out.push({
					label: folder.name,
					hint: folder.unread ? `${label} — ${folder.unread} unread` : label,
					terms: "folder go to",
					run: () => go(a.id, folder.id, folder.name),
				});
			}
			for (const c of this.plugin.categoriesFor(a.id)) {
				out.push({
					label: c.displayName,
					hint: `${label} — category`,
					terms: "category tag label",
					run: () => go(a.id, categoryFolderId(c.displayName), c.displayName),
				});
			}
			// a search folder is a saved query, not a place: it runs rather
			// than navigates, exactly as clicking it in the tree does
			for (const sf of s.mailSearchFolders.filter((f) => f.accountId === a.id)) {
				out.push({
					label: sf.name,
					hint: `${label} — search folder`,
					terms: "search folder saved",
					run: () => {
						this.screen = "list";
						if (this.searchInputEl) this.searchInputEl.value = sf.query;
						void this.plugin.runMailSearch(sf.query);
						this.render();
					},
				});
			}
		}

		out.push({ label: "New mail", hint: "C", terms: "compose write", run: () => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open() });
		out.push({ label: "Search", hint: "/", terms: "find from subject is unread flagged attachment phrase", run: () => this.focusSearch() });
		out.push({ label: "Advanced search", terms: "find fields date range from subject", run: () => this.openSearchWindow() });
		out.push({ label: "People", terms: "contacts address book who", run: () => this.plugin.openPeople() });
		out.push({ label: "Tasks", terms: "todo to-do flagged", run: () => this.plugin.openTasks() });
		out.push({ label: "Outlook notes", terms: "sticky notes", run: () => new NotesModal(this.app, this.plugin).open() });
		out.push({ label: "Journal", terms: "day activity diary daily note", run: () => new JournalModal(this.app, this.plugin).open() });
		out.push({ label: "Folders", terms: "manage tree mark all read hidden", run: () => this.plugin.openFolders() });
		out.push({ label: "Shortcuts", terms: "favourites favorites jump launcher", run: () => this.plugin.openShortcuts() });
		out.push({ label: "Refresh", terms: "sync fetch", run: () => this.plugin.userRefreshMail() });
		if (Platform.isDesktopApp) out.push({ label: "Print the list", hint: `${this.lastList.length} shown`, terms: "table paper pdf", run: () => this.printList() });

		// actions that need something open only exist when something is
		const m = this.selected;
		if (m) {
			const n = this.multiTargets().length;
			const many = n > 1 ? ` (${n})` : "";
			out.push({ label: "Reply", hint: "R", terms: "answer", run: () => this.composeFor("reply") });
			out.push({ label: "Reply all", hint: "A", terms: "answer everyone", run: () => this.composeFor("replyAll") });
			out.push({ label: "Forward", hint: "F", terms: "send on", run: () => this.composeFor("forward") });
			out.push({ label: `Archive${many}`, hint: "E", terms: "file away", run: () => this.archiveSelection() });
			out.push({ label: `Delete${many}`, hint: "Del", terms: "bin trash", run: () => this.deleteSelection(false) });
			out.push({ label: `Delete${many} permanently`, hint: "Shift + Del", terms: "bin trash purge", run: () => this.deleteSelection(true) });
			out.push({ label: `Mark${many} read`, hint: "Shift + I", terms: "seen", run: () => this.markSelection(true) });
			out.push({ label: `Mark${many} unread`, hint: "Shift + U", terms: "seen", run: () => this.markSelection(false) });
			out.push({ label: this.multiTargets().every((x) => x.flagged) ? `Clear the flag${many}` : `Flag${many}`, hint: "S", terms: "follow up star", run: () => this.flagSelection() });
			out.push({ label: `Categorize${many}...`, hint: "G", terms: "tag label color", run: () => this.categorizeMenu(this.multiTargets()) });
			out.push({ label: `Snooze${many}`, hint: "B", terms: "later remind defer", run: () => this.snoozeSelection() });
			out.push({ label: `Move${many} to...`, hint: "V", terms: "file folder", run: () => this.openMovePicker(this.multiTargets()) });
			for (const r of s.mailRecentFolders.slice(0, 5)) {
				if (!this.multiTargets().some((t) => t.accountId === r.accountId)) continue;
				out.push({
					label: `Move${many} to ${r.name}`,
					hint: "recent",
					terms: "file folder",
					run: () => {
						const targets = this.multiTargets();
						void this.plugin.moveMail(targets, r.accountId, r.folderId, r.name).then(() => {
							if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
							this.multiSel.clear();
							this.render();
						});
					},
				});
			}
			if (Platform.isDesktopApp) out.push({ label: "Print", terms: "paper pdf", run: () => void this.printMessage(m) });
			out.push({ label: "Make event", terms: "calendar meeting", run: () => this.makeEventFrom(m) });
			out.push({ label: "Save to note", terms: "vault capture", run: () => void this.plugin.saveMailToNote(m) });
			if (m.webLink) out.push({ label: "Open in Outlook", terms: "browser web", run: () => window.open(m.webLink, "_blank") });
			const t = this.selectedThread();
			if (t) out.push({ label: this.selectedThreadOpen() ? "Collapse the conversation" : "Expand the conversation", hint: "O", terms: "thread", run: () => this.setThreadOpen(!this.selectedThreadOpen()) });
		}

		const toggle = (label: string, on: boolean, run: () => void, terms?: string) => out.push({ label, hint: on ? "on" : "off", terms, run });
		toggle("Split the inbox into sections", s.mailSplitInbox, () => {
			s.mailSplitInbox = !s.mailSplitInbox;
			this.plugin.queueSave();
			this.render();
		}, "bundle sections focused notifications");
		toggle("Show as conversations", s.mailConversations, () => {
			s.mailConversations = !s.mailConversations;
			this.plugin.queueSave();
			this.render();
		}, "thread group");
		toggle("Unread filter", s.mailUnreadOnly, () => {
			s.mailUnreadOnly = !s.mailUnreadOnly;
			this.plugin.queueSave();
			this.render();
		}, "filter");
		for (const d of ["compact", "cozy", "comfortable"] as const) {
			out.push({
				label: `Density: ${d}`,
				hint: s.mailDensity === d ? "current" : undefined,
				terms: "spacing rows",
				run: () => {
					s.mailDensity = d;
					this.plugin.queueSave();
					this.render();
				},
			});
		}
		for (const n of [0, 1, 2, 3] as const) {
			out.push({
				label: `Message preview: ${n === 0 ? "off" : `${n} line${n === 1 ? "" : "s"}`}`,
				hint: s.mailPreviewLines === n ? "current" : undefined,
				terms: "preview lines density",
				run: () => {
					s.mailPreviewLines = n;
					this.plugin.queueSave();
					this.render();
				},
			});
		}
		const waiting = this.plugin.snoozedFor();
		if (waiting.length) {
			out.push({
				label: `Snoozed mail (${waiting.length})`,
				hint: `next ${fmtWhen(waiting[0].dueMs, s.use24h)}`,
				terms: "later waiting deferred",
				run: () => new SnoozedModal(this.app, this.plugin, () => this.render()).open(),
			});
		}
		out.push({ label: "Manage categories", terms: "tag label color new delete", run: () => new CategoriesModal(this.app, this.plugin, () => this.render()).open() });
		out.push({ label: "Inbox rules", terms: "filter automate server", run: () => new RulesModal(this.app, this.plugin).open() });
		{
			const on = this.plugin.autoReplyingAccounts();
			out.push({
				label: "Automatic replies (out of office)",
				hint: on.length ? "on" : "off",
				terms: "vacation away holiday ooo",
				run: () => new OutOfOfficeModal(this.app, this.plugin).open(),
			});
		}
		if (m) {
			out.push({
				label: "Create a rule from this message",
				terms: "filter automate",
				run: () => new RuleEditModal(this.app, this.plugin, m.accountId, null, { name: `From ${m.from}`, fromContains: m.fromAddress || m.from }, () => this.render()).open(),
			});
		}
		out.push({ label: "Keyboard shortcuts", hint: "?", terms: "help keys", run: () => new MailShortcutsModal(this.app).open() });
		return out;
	}

	/** Move the reading selection through the list exactly as rendered, so
	 *  an expanded conversation's replies are stops on the way down and a
	 *  collapsed one is a single stop. */
	private moveSelection(n: number) {
		if (!this.lastList.length) return;
		const i = this.selected ? this.lastList.findIndex((x) => x.id === this.selected?.id) : -1;
		const next = i < 0 ? (n > 0 ? 0 : this.lastList.length - 1) : Math.min(this.lastList.length - 1, Math.max(0, i + n));
		const m = this.lastList[next];
		if (!m || m.id === this.selected?.id) return;
		this.multiSel.clear();
		this.selAnchorId = m.id;
		void this.select(m);
		this.scrollSelectedIntoView();
	}

	/** Keep the keyboard's selection on screen. Runs after the render that
	 *  the selection triggered, which is why it waits a tick. */
	private scrollSelectedIntoView() {
		window.setTimeout(() => {
			const i = this.lastList.findIndex((x) => x.id === this.selected?.id);
			const rows = this.listEl?.querySelectorAll(".pcal-mail-row");
			if (i >= 0 && rows?.[i]) (rows[i] as HTMLElement).scrollIntoView({ block: "nearest" });
		}, 0);
	}

	/** Enter reads the selection; on a narrow pane that means drilling in. */
	private openSelected() {
		if (!this.selected) {
			if (this.lastList.length) void this.select(this.lastList[0]);
			return;
		}
		if (this.drill) {
			this.screen = "read";
			this.render();
		}
	}

	/** u backs out: out of the message on a narrow pane, otherwise out of a
	 *  bulk selection. */
	private goBackOrClear() {
		if (this.drill && this.screen === "read") {
			this.screen = "list";
			this.render();
			return;
		}
		if (this.multiSel.size) {
			this.multiSel.clear();
			this.render();
		}
	}

	private escape() {
		if (this.searchRowEl?.isShown()) {
			this.searchInputEl.value = "";
			this.searchRowEl.hide();
			this.plugin.clearMailSearch();
		}
		if (this.multiSel.size || this.selectMode) {
			this.multiSel.clear();
			this.selectMode = false;
		}
		this.render();
	}

	private focusSearch() {
		if (!this.searchRowEl) return;
		this.searchRowEl.show();
		this.searchInputEl.focus();
		this.searchInputEl.select();
	}

	/** Run a query from the search window, in the box so it can be edited by
	 *  hand afterwards rather than being a search you cannot see. */
	runSearch(q: string, everywhere: boolean) {
		this.screen = "list";
		this.searchRowEl?.show();
		if (this.searchInputEl) this.searchInputEl.value = q;
		if (everywhere) void this.plugin.runMailSearch(q);
		else this.plugin.searchLocal(q);
		this.render();
	}

	/** Show one folder, from somewhere outside the folder pane. */
	goFolder(accountId: string, folderId: string, name: string) {
		this.plugin.clearMailSearch();
		this.folderSel = { accountId, folderId, name };
		this.screen = "list";
		this.autoSelectPending = true;
		this.render();
	}

	openSearchWindow() {
		new SearchModal(this.app, this.plugin, this, this.searchInputEl?.value).open();
	}

	private archiveSelection() {
		const targets = this.multiTargets();
		if (!targets.length) return;
		for (const t of targets) void this.plugin.archiveMail(t);
		if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
		this.multiSel.clear();
		this.render();
	}

	private markSelection(read: boolean) {
		const targets = this.multiTargets();
		for (const t of targets) void this.plugin.setMailRead(t, read);
		this.render();
	}

	/** S toggles the flag. A mixed set flags rather than clears, so pressing
	 *  it once over a selection always ends with everything flagged and the
	 *  second press clears the lot. */
	private flagSelection() {
		const targets = this.multiTargets();
		if (!targets.length) return;
		const on = !targets.every((t) => t.flagged);
		for (const t of targets) void this.plugin.setMailFlag(t, on);
		this.render();
	}

	/** An email as a prefilled event at the next half hour, sender invited. */
	private makeEventFrom(m: PCMail) {
		const now = Date.now();
		const min = Math.ceil((minutesOfMs(now) + 1) / 30) * 30;
		const dayKey = min >= 1440 ? addDays(keyOfMs(now), 1) : keyOfMs(now);
		const start = msOfKey(dayKey) + (min % 1440) * 60000;
		new EventModal(this.app, this.plugin, null, start, start + 30 * 60000, false, { title: subjectToEventTitle(m.subject), invites: m.fromAddress }).open();
	}

	private snoozeSelection() {
		this.snoozeTargets(this.multiTargets());
	}

	/** Open the folder picker for a set of messages. */
	private openMovePicker(targets: PCMail[]) {
		if (!targets.length) return;
		new MoveToFolderModal(this.app, this.plugin, targets, () => {
			if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
			this.multiSel.clear();
			this.render();
		}).open();
	}

	/** Park messages until a time, and take them out of the list now. */
	private snoozeTargets(targets: PCMail[]) {
		if (!targets.length) return;
		const what = targets.length > 1 ? `${targets.length} messages` : `"${targets[0].subject}"`;
		new WhenModal(
			this.app,
			targets.length > 1 ? `Snooze ${targets.length} messages` : "Snooze",
			"Snoozed mail waits in a Snoozed folder in your mailbox and comes back to the inbox at this time, unread. The return needs Obsidian to be running, so if it is closed the message comes back the next time you open it.",
			this.plugin.settings.use24h,
			(ms) => {
				void (async () => {
					let ok = 0;
					for (const t of targets) if (await this.plugin.snoozeMail(t, ms)) ok++;
					if (ok) new Notice(`Power Desk: ${what} back ${fmtWhen(ms, this.plugin.settings.use24h)}.`);
					if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
					this.multiSel.clear();
					this.render();
				})();
			}
		).open();
	}

	private composeFor(mode: "reply" | "replyAll" | "forward") {
		if (!this.selected) return;
		new RichComposeModal(this.app, this.plugin, { mode, mail: this.selected }).open();
	}

	/** x ticks the row under the reading selection, the way Gmail's does. */
	private toggleSelectionMark() {
		const m = this.selected;
		if (!m) return;
		const i = this.lastList.findIndex((x) => x.id === m.id);
		const targets = i >= 0 ? this.lastRows[i] : [m];
		this.setSelected(targets, !targets.every((x) => this.multiSel.has(x.id)));
		this.selAnchorId = m.id;
		this.render();
	}

	/** The conversation the reading selection sits in, when it is one. */
	private selectedThread(): MailThread | null {
		const m = this.selected;
		if (!m) return null;
		return this.lastThreads.find((t) => t.messages.length > 1 && t.messages.some((x) => x.id === m.id)) ?? null;
	}

	private selectedThreadOpen(): boolean {
		const t = this.selectedThread();
		return !!t && this.expanded.has(t.key);
	}

	private setThreadOpen(open: boolean) {
		const t = this.selectedThread();
		if (!t) return;
		if (open) this.expanded.add(t.key);
		else this.expanded.delete(t.key);
		this.render();
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
		// the columns below are built fresh, so whatever the last open drew is
		// gone; a stale signature would let the first render skip itself
		this.drawnItems.clear();
		this.lastReadSig = null;
		this.lastFolderSig = null;
		root.addClass("pcal-mail-root");
		const header = root.createDiv("pcal-mail-header");
		this.backBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Back" } });
		setIcon(this.backBtn, "chevron-left");
		this.backBtn.addEventListener("click", () => {
			this.screen = this.screen === "read" ? "list" : "folders";
			this.render();
		});
		// this button belongs to the vault's page tree, not to the mail folder
		// pane: the folders here are the point of the view and are never the
		// thing you want out of the way
		// "files" rather than a panel glyph: this toggles the vault's notes,
		// not a panel belonging to this view, and the calendar has a real
		// panel toggle of its own that should keep the panel icon
		this.foldToggleBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Full width: fold the vault's notes away while this tab is open" } });
		setIcon(this.foldToggleBtn, "panel-left");
		this.foldToggleBtn.toggleClass("is-active", this.plugin.focusOn());
		this.foldToggleBtn.addEventListener("click", () => {
			this.plugin.toggleFocus();
			this.syncFoldToggle();
		});
		this.register(this.plugin.watchFocus(() => this.syncFoldToggle()));
		this.toolsEl = header.createDiv("pcal-mail-tools");
		this.renderToolbar();
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
		// the box for a word, the window for a question with several parts
		searchBtn.addEventListener("contextmenu", (e) => {
			e.preventDefault();
			const menu = new Menu();
			menu.addItem((i) => i.setTitle("Advanced search...").setIcon("search").onClick(() => this.openSearchWindow()));
			menu.showAtMouseEvent(e);
		});
		const advBtn = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Advanced search (Shift + /)" } });
		setIcon(advBtn, pickIcon("search-check", "filter", "search"));
		advBtn.addEventListener("click", () => this.openSearchWindow());
		this.refreshBtn = right.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Refresh" } });
		setIcon(this.refreshBtn, "refresh-cw");
		this.refreshBtn.addEventListener("click", () => this.plugin.userRefreshMail());
		this.oooEl = root.createDiv("pcal-mail-ooo-bar");
		this.oooEl.hide();
		this.searchRowEl = root.createDiv("pcal-mail-searchrow");
		this.searchRowEl.hide();
		this.searchInputEl = this.searchRowEl.createEl("input", { attr: { type: "search", placeholder: "Search: words, from:name, subject:x, is:unread, has:attachment, \"a phrase\"" } });
		// typing searches the local index, which is instant and costs nothing;
		// Enter widens the same query to the whole mailbox
		this.searchInputEl.addEventListener("input", () => {
			if (this.searchTypeTimer != null) window.clearTimeout(this.searchTypeTimer);
			this.searchTypeTimer = window.setTimeout(() => {
				this.searchTypeTimer = null;
				this.plugin.searchLocal(this.searchInputEl.value);
			}, 90);
		});
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
			menu.addSeparator();
			menu.addItem((i) =>
				i
					.setTitle("Show as conversations")
					.setChecked(this.plugin.settings.mailConversations)
					.onClick(() => {
						this.plugin.settings.mailConversations = !this.plugin.settings.mailConversations;
						this.plugin.queueSave();
						this.render();
					})
			);
			menu.addItem((i) =>
				i
					.setTitle("Split the inbox into sections")
					.setChecked(this.plugin.settings.mailSplitInbox)
					.onClick(() => {
						this.plugin.settings.mailSplitInbox = !this.plugin.settings.mailSplitInbox;
						this.plugin.queueSave();
						this.render();
					})
			);
			if (Platform.isDesktopApp) {
				menu.addSeparator();
				menu.addItem((i) => i.setTitle("Print this list...").setIcon("printer").onClick(() => this.printList()));
			}
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
				pane.addClass("pcal-pane-sized");
				pane.style.width = `${saved}px`;
			}
			el.addEventListener("pointerdown", (e) => {
				e.preventDefault();
				const startX = e.clientX;
				const startW = pane.getBoundingClientRect().width;
				pane.addClass("pcal-pane-sized");
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
	/** Everything the toolbar can carry.
	 *
	 *  `many` says whether an action reads the whole bulk selection or wants
	 *  one message: replying to nine at once is not a thing, deleting nine is.
	 *  Ordering and membership are the user's, so this is a catalog rather
	 *  than a layout. Nothing here is only here: every action also has a
	 *  shortcut, a right-click entry, or a palette line, so a toolbar trimmed
	 *  to four buttons loses no capability. */
	/** `solo` marks an action that needs no message at all: it acts on the list
	 *  or opens a window of its own, so demanding a selection first would only
	 *  be a rule with nothing behind it. */
	private toolActions(): { id: string; label: string; icon: string; many: boolean; solo?: boolean; run: (targets: PCMail[]) => void }[] {
		const one = (targets: PCMail[]): PCMail | null => {
			if (targets.length > 1) {
				new Notice("Power Desk: this works on a single message; pick just one.");
				return null;
			}
			return targets[0] ?? null;
		};
		return [
			{
				id: "delete",
				label: "Delete",
				icon: "trash-2",
				many: true,
				run: (t) => this.deleteMessages(t, false),
			},
			{
				id: "report",
				label: "Report",
				icon: "shield-alert",
				many: true,
				run: (t) => this.reportMenu(t),
			},
			{ id: "reply", label: "Reply", icon: "reply", many: false, run: (t) => { const m = one(t); if (m) void this.selectAndReply(m); } },
			{ id: "replyAll", label: "Reply all", icon: "reply-all", many: false, run: (t) => { const m = one(t); if (m) void this.selectAndReply(m, true); } },
			{
				id: "forward",
				label: "Forward",
				icon: "forward",
				many: false,
				run: (t) => {
					const m = one(t);
					if (m) new RichComposeModal(this.app, this.plugin, { mode: "forward", mail: m }).open();
				},
			},
			{
				id: "readUnread",
				label: "Read / Unread",
				icon: "mail-open",
				many: true,
				run: (t) => {
					const anyUnread = t.some((x) => x.unread);
					for (const x of t) void this.plugin.setMailRead(x, anyUnread);
					this.render();
				},
			},
			{
				id: "flag",
				label: "Flag / Unflag",
				icon: "flag",
				many: true,
				run: (t) => {
					const on = !t.every((x) => x.flagged);
					for (const x of t) void this.plugin.setMailFlag(x, on);
					this.render();
				},
			},
			{ id: "categorize", label: "Categorize", icon: "tag", many: true, run: (t) => this.categorizeMenu(t) },
			{
				id: "print",
				label: "Print",
				icon: "printer",
				many: true,
				solo: true,
				// the window carries both styles, so one message opens on the
				// memo and a selection of several (or none) opens on the table
				run: (t) => void this.openPrintWindow(t.length === 1 ? "memo" : "table", t.length === 1 ? t[0] : null),
			},
			{ id: "printList", label: "Print the list", icon: "table", many: true, solo: true, run: () => this.printList() },
			{ id: "people", label: "People", icon: "users", many: true, run: () => this.plugin.openPeople() },
			{ id: "tasks", label: "Tasks", icon: "check-square", many: true, run: () => this.plugin.openTasks() },
			{ id: "notes", label: "Notes", icon: pickIcon("sticky-note", "file-text"), many: true, run: () => new NotesModal(this.app, this.plugin).open() },
			{ id: "journal", label: "Journal", icon: pickIcon("book-open", "book", "file-text"), many: true, run: () => new JournalModal(this.app, this.plugin).open() },
			{ id: "folders", label: "Folders", icon: "folder", many: true, run: () => this.plugin.openFolders() },
			{ id: "shortcuts", label: "Shortcuts", icon: "bookmark", many: true, run: () => this.plugin.openShortcuts() },
			{ id: "snooze", label: "Snooze", icon: "clock", many: true, run: (t) => this.snoozeTargets(t) },
			{ id: "move", label: "Move", icon: "folder-input", many: true, run: (t) => this.openMovePicker(t) },
			{
				id: "archive",
				label: "Archive",
				icon: "archive",
				many: true,
				run: (t) => {
					for (const x of t) void this.plugin.archiveMail(x);
					if (t.some((x) => x.id === this.selected?.id)) this.selected = null;
					this.multiSel.clear();
					this.render();
				},
			},
			{
				id: "makeEvent",
				label: "Make event",
				icon: "calendar-plus",
				many: false,
				run: (t) => {
					const m = one(t);
					if (m) this.makeEventFrom(m);
				},
			},
			{
				id: "saveToNote",
				label: "Save to note",
				icon: "file-plus",
				many: false,
				run: (t) => {
					const m = one(t);
					if (m) void this.plugin.saveMailToNote(m);
				},
			},
			{
				id: "rule",
				label: "Create a rule",
				icon: "filter",
				many: false,
				run: (t) => {
					const m = one(t);
					if (m) new RuleEditModal(this.app, this.plugin, m.accountId, null, { name: `From ${m.from}`, fromContains: m.fromAddress || m.from }, () => this.render()).open();
				},
			},
			{
				id: "openInOutlook",
				label: "Open in Outlook",
				icon: "external-link",
				many: false,
				run: (t) => {
					const m = one(t);
					if (m?.webLink) window.open(m.webLink, "_blank");
				},
			},
		];
	}

	/** Draw the toolbar from the saved order. Rebuilt whenever the order
	 *  changes; the buttons themselves enable and disable per render. */
	private renderToolbar() {
		const host = this.toolsEl;
		if (!host) return;
		host.empty();
		this.mailToolBtns = [];
		const newMail = host.createEl("button", { cls: "pcal-new-btn", text: "New mail" });
		newMail.addEventListener("click", () => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open());
		const catalog = new Map(this.toolActions().map((a) => [a.id, a]));
		for (const id of this.plugin.settings.mailToolbar) {
			const a = catalog.get(id);
			if (!a) continue;
			const b = host.createEl("button", { cls: "pcal-icon-btn pcal-mail-tool", attr: { "aria-label": a.label } });
			setIcon(b.createSpan(), a.icon);
			b.createSpan({ cls: "pcal-mail-tool-label", text: a.label });
			b.addEventListener("click", () => {
				const targets = this.multiTargets();
				if (!targets.length && !a.solo) {
					new Notice("Power Desk: select a message first.");
					return;
				}
				a.run(targets);
			});
			if (!a.solo) this.mailToolBtns.push(b);
		}
		const more = host.createEl("button", { cls: "pcal-icon-btn pcal-mail-tool-more", attr: { "aria-label": "Customize the toolbar" } });
		setIcon(more, "settings-2");
		more.addEventListener("click", () =>
			new ToolbarModal(this.app, this.plugin, this.toolActions(), { get: () => this.plugin.settings.mailToolbar, set: (v) => (this.plugin.settings.mailToolbar = v), fallback: DEFAULT_SETTINGS.mailToolbar, leads: "New mail" }, () => {
				this.renderToolbar();
				this.render();
			}).open()
		);
	}

	openToolbarEditor() {
		new ToolbarModal(this.app, this.plugin, this.toolActions(), { get: () => this.plugin.settings.mailToolbar, set: (v) => (this.plugin.settings.mailToolbar = v), fallback: DEFAULT_SETTINGS.mailToolbar, leads: "New mail" }, () => {
			this.renderToolbar();
			this.render();
		}).open();
	}

	/** The mailbox's categories as a checklist, applied to the selection.
	 *
	 *  Ticked means every message in the selection carries it, so toggling
	 *  over a mixed set puts it on all of them first, and the second press
	 *  takes it off all of them. Same rule as the flag, for the same reason:
	 *  one press should have a predictable end state. */
	private categorizeMenu(targets: PCMail[]) {
		const accountId = targets[0]?.accountId;
		if (!accountId) return;
		void this.plugin.loadCategories(accountId).then((cats) => {
			const menu = new Menu();
			const mine = targets.filter((t) => t.accountId === accountId);
			if (!cats.length) menu.addItem((i) => i.setTitle("No categories in this mailbox yet").setDisabled(true));
			for (const c of cats) {
				const all = mine.every((t) => (t.categories ?? []).some((x) => x.toLowerCase() === c.displayName.toLowerCase()));
				menu.addItem((i) =>
					i
						.setTitle(c.displayName)
						.setChecked(all)
						.onClick(() => {
							// ticked everywhere means take it off; anything else
							// means put it on, so one press has one end state
							const want = !all;
							for (const t of mine) {
								const has = (t.categories ?? []).some((x) => x.toLowerCase() === c.displayName.toLowerCase());
								if (has === want) continue;
								void this.plugin.setMailCategories(t, toggleCategory(t.categories, c.displayName));
							}
							this.render();
						})
				);
			}
			if (cats.length && mine.some((t) => (t.categories ?? []).length)) {
				menu.addSeparator();
				menu.addItem((i) =>
					i
						.setTitle("Clear all categories")
						.setIcon("x")
						.onClick(() => {
							for (const t of mine) void this.plugin.setMailCategories(t, []);
							this.render();
						})
				);
			}
			menu.addSeparator();
			menu.addItem((i) =>
				i
					.setTitle("Manage categories...")
					.setIcon("settings-2")
					.onClick(() => new CategoriesModal(this.app, this.plugin, () => this.render()).open())
			);
			const rect = this.toolsEl?.getBoundingClientRect();
			if (rect) menu.showAtPosition({ x: rect.left, y: rect.bottom + 2 });
		});
	}

	/** Report as junk, and an honest word about what phishing reporting
	 *  actually is. */
	private reportMenu(targets: PCMail[]) {
		const menu = new Menu();
		menu.addItem((i) =>
			i
				.setTitle(targets.length > 1 ? `Report ${targets.length} as junk` : "Report as junk")
				.setIcon("ban")
				.onClick(() => {
					new ConfirmModal(
						this.app,
						targets.length > 1 ? `Report ${targets.length} messages as junk?` : `Report as junk?`,
						"The sender goes on your blocked list and the mail moves to Junk Email, which is what teaches the filter.",
						"Report junk",
						() => {
							void this.plugin.reportJunk(targets).then((n) => {
								if (n) new Notice(`Power Desk: reported ${n} message${n === 1 ? "" : "s"} as junk.`);
								if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
								this.multiSel.clear();
								this.render();
							});
						}
					).open();
				})
		);
		const n = targets.length;
		const preview = (action: "phish" | "notJunk", title: string, heading: string, body: string) =>
			menu.addItem((i) =>
				i
					.setTitle(title)
					.setIcon(action === "phish" ? "shield-alert" : "mail-check")
					.onClick(() => {
						new ConfirmModal(this.app, heading, body, title, () => {
							void this.plugin.reportPreview(targets, action).then((r) => {
								if (r.done) new Notice(action === "phish" ? `Power Desk: reported ${r.done} message${r.done === 1 ? "" : "s"} as phishing.` : `Power Desk: told your mailbox ${r.done} message${r.done === 1 ? "" : "s"} were not junk.`);
								if (r.error)
									new Notice(
										`Power Desk: that report did not go through. ${r.error} This one runs on Microsoft's preview API, which they can change without notice; Report as junk uses the supported one and still works.`,
										14000
									);
								if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
								this.multiSel.clear();
								this.render();
							});
						}).open();
					})
			);
		menu.addSeparator();
		preview(
			"phish",
			n > 1 ? `Report ${n} as phishing (preview)` : "Report phishing (preview)",
			n > 1 ? `Report ${n} messages as phishing?` : "Report as phishing?",
			"The message goes to Microsoft's filters as a phishing report and moves out of the inbox. This runs on Microsoft's preview API, which they state is not supported for production and may change without notice; if it stops working, Outlook's own Report button always will."
		);
		preview(
			"notJunk",
			n > 1 ? `Report ${n} as not junk (preview)` : "Report as not junk (preview)",
			n > 1 ? `Report ${n} messages as not junk?` : "Report as not junk?",
			"This tells your mailbox the message was filed wrongly and moves it back to the inbox. It runs on the same preview API as the phishing report, because the supported version of this one was retired."
		);
		const rect = this.toolsEl?.getBoundingClientRect();
		if (rect) menu.showAtPosition({ x: rect.left, y: rect.bottom + 2 });
	}

	/** Print one message.
	 *
	 *  Through a hidden iframe carrying its own document, because asking the
	 *  window to print prints Obsidian: the sidebar, the folder tree, the
	 *  list, and whatever slice of the message fits beside them. The iframe
	 *  is its own page with its own stylesheet, so what comes out is the mail
	 *  and nothing else. */
	private async messageStyle(m: PCMail): Promise<PrintStyle> {
		const body = (await this.plugin.readMailBody(m)) ?? this.selectedBody;
		let html = body?.html ?? "";
		const cids = this.selected?.id === m.id ? this.inlineCids : null;
		if (html && cids) for (const [cid, url] of cids) html = html.split(`cid:${cid}`).join(url);
		// the same sanitizer the reading pane uses, so a printed message can
		// no more run something than a read one can
		let safe = "";
		if (html) {
			const holder = createDiv();
			holder.appendChild(sanitizeHTMLToDom(html));
			safe = holder.innerHTML;
		}
		const head = {
			subject: m.subject,
			from: `${m.from}${m.fromAddress ? ` <${m.fromAddress}>` : ""}`,
			to: body?.toLine ?? "",
			date: `${fmtDayShort(keyOfMs(m.receivedMs), true)} ${fmtTimeOfMs(m.receivedMs, this.plugin.settings.use24h)}`,
			bodyHtml: safe || body?.text || "",
			plain: !safe,
		};
		// fetched once and closed over: changing the size must not go back to
		// the server for a body that has not changed
		return { id: "memo", label: "Memo", hint: "This message", icon: "mail", basePt: 12, landscape: false, build: (o) => printableHtml(head, o) };
	}

	/** The list as a table, Outlook's other print style. Prints exactly what
	 *  is on screen, filters, sort, and all, because a printed list that
	 *  quietly differs from the one you were looking at is worse than none. */
	private listStyle(): PrintStyle | null {
		if (!this.lastList.length) return null;
		const s = this.plugin.settings;
		const todayKey = keyOfDate(new Date());
		const names = new Map(this.plugin.mailAccounts().map((a) => [a.id, this.plugin.folderNamesFor(a.id)]));
		const search = this.plugin.mailSearchState();
		const title = search ? `Search: ${search.query}` : this.folderSel ? this.folderSel.name : "All inboxes";
		const rows = this.lastList.map((m) => ({
			from: m.from,
			subject: m.subject,
			date: fmtMailTime(m.receivedMs, todayKey, s.use24h),
			folder: (m.folderId ? names.get(m.accountId)?.get(m.folderId) : "") || "",
		}));
		return {
			id: "table",
			label: "Table",
			hint: `${rows.length} shown`,
			icon: "table",
			basePt: 10,
			landscape: false,
			build: (o) => printableTableHtml(title, rows, o),
		};
	}

	/** Both mail styles in one window, whichever was asked for on top. Outlook
	 *  offers Memo and Table from the same dialog, and the choice between them
	 *  is exactly the sort of thing you make by looking rather than by knowing
	 *  in advance which menu item you wanted. */
	private async openPrintWindow(prefer: "memo" | "table", m?: PCMail | null) {
		if (!Platform.isDesktopApp) {
			new Notice("Power Desk: printing needs the desktop app.");
			return;
		}
		const styles: PrintStyle[] = [];
		const one = m ?? this.selected;
		if (one) styles.push(await this.messageStyle(one));
		const list = this.listStyle();
		if (list) styles.push(list);
		if (!styles.length) {
			new Notice("Power Desk: open a message or a list first, so there is something to print.");
			return;
		}
		new PrintModal(this.app, styles, prefer).open();
	}

	private printMessage(m: PCMail) {
		void this.openPrintWindow("memo", m);
	}

	private printList() {
		void this.openPrintWindow("table");
	}

	/** Jump to one message, from a notification. Leaves any folder or search
	 *  and goes to the unified list, which is the one place every account's
	 *  mail is certain to be. */
	revealMessage(m: PCMail) {
		this.plugin.clearMailSearch();
		this.folderSel = null;
		this.screen = this.drill ? "read" : "list";
		this.multiSel.clear();
		void this.select(m);
		this.scrollSelectedIntoView();
	}

	/** The folder this view is showing, refreshed on the mail poll. A search
	 *  is left alone: its results answer a query rather than being a live
	 *  list, and rerunning it underneath would move rows while they are read. */
	pollCurrentFolder() {
		if (this.plugin.mailSearchState()) return;
		if (this.folderSel) this.plugin.ensureFolderMail(this.folderSel.accountId, this.folderSel.folderId, true);
	}

	/** Which list was on screen last render, so a change of list can be seen. */
	private lastListKey: string | null = null;

	private autoReadSuppressed(): boolean {
		const s = this.plugin.settings;
		return (s.mailUnreadOnly || this.folderSel?.folderId === UNREAD_FOLDER) && s.unreadFilterKeepsUnread;
	}

	private queueRender() {
		if (this.renderQueued) return;
		this.renderQueued = true;
		window.requestAnimationFrame(() => {
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
		const accounts = this.plugin.mailAccounts();
		if (!accounts.length) {
			this.listEl.empty();
			this.drawnItems.clear();
			const empty = this.listEl.createDiv("pcal-empty");
			empty.createDiv({ text: "No mail-enabled account. Reconnect a Microsoft account in settings to grant mail access." });
			const b = empty.createEl("button", { text: "Open settings", cls: "mod-cta" });
			b.addEventListener("click", () => this.plugin.openOwnSettings());
			this.readEl.empty();
			return;
		}
		const colorOf = new Map(accounts.map((a, i) => [a.id, paletteColor(i)]));
		this.renderOooBar();
		this.renderFolders(accounts, colorOf);
		const sel = this.folderSel;
		if (sel && !this.plugin.accountById(sel.accountId)) this.folderSel = null;
		const search = this.plugin.mailSearchState();
		// leaving a list ends the reprieve its rows were given: the exemption
		// exists so a row does not vanish under your hand, not so a read
		// message follows you into every list you open afterwards
		const listKey = search ? `search:${search.query}` : this.folderSel ? `${this.folderSel.accountId}:${this.folderSel.folderId}` : "all";
		if (listKey !== this.lastListKey) {
			if (this.lastListKey != null) this.plugin.recentlyMarked.clear();
			this.lastListKey = listKey;
		}
		// the title says what was actually searched, because "5 results" over
		// a local index and over the whole mailbox are different claims
		const scopeLabel = search ? (search.inFlight ? "searching the mailbox" : search.scope === "local" ? `${search.results.length} here` : `${search.results.length} in the mailbox`) : "";
		this.titleTextEl.setText(
			search ? `${search.query} (${scopeLabel})` : this.drill && this.screen === "folders" ? "Mailboxes" : this.folderSel ? this.folderSel.name : "All inboxes"
		);
		this.titleTextEl.setAttribute("aria-label", search && search.scope === "local" ? "Results from mail already on this device; press Enter to search the whole mailbox" : "");
		if (!search && this.folderSel) this.plugin.ensureFolderMail(this.folderSel.accountId, this.folderSel.folderId, false);
		const source = search ? search.results : this.folderSel ? this.plugin.folderMail(this.folderSel.accountId, this.folderSel.folderId) : this.plugin.allMail();
		// the unread filter keeps the currently open message visible even once
		// it reads, so it cannot vanish mid-read
		const mail = source.filter((m) => !s.mailUnreadOnly || m.unread || m.id === this.selected?.id || this.plugin.recentlyMarked.has(m.id));
		// a screenful and then some, not eight: one $batch carries twenty, so
		// covering what you can see costs about the same as covering a third
		// of it, and the ninth message down was the one that felt slow
		void this.plugin.prefetchBodies(mail.slice(0, 25));
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
		// errors and the empty line are worked out here but drawn with the rows,
		// so the whole list is one comparison and one rebuild rather than a
		// half-emptied column whenever a refresh changes nothing
		const errors: string[] = search?.error
			? [search.error]
			: this.folderSel && !search
				? [this.plugin.folderMailError(this.folderSel.accountId, this.folderSel.folderId)].filter((e): e is string => !!e)
				: search
					? []
					: this.plugin.mailErrors();
		const emptyText = mail.length
			? null
			: search
				? search.inFlight
					? "Searching the mailbox..."
					: search.scope === "local"
						? "Nothing here yet. Press Enter to search the whole mailbox."
						: "No matches."
				: this.plugin.anyMailInFlight()
					? "Loading..."
					: s.mailUnreadOnly
						? "No unread mail."
						: this.folderSel
							? "Nothing here."
							: "Inbox zero.";
		// search results, the Unread folder, and a category all gather mail from
		// anywhere in the mailbox, so their rows say which folder it came from
		const showTags = !!search || this.folderSel?.folderId === UNREAD_FOLDER || !!folderIdCategory(this.folderSel?.folderId ?? "");
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
		// The split inbox groups the triage surfaces into sections; browsing
		// a filing folder or Sent Items stays flat, where splitting would
		// only be noise, and a search always does.
		const triage = !search && (!this.folderSel || this.folderSel.folderId === UNREAD_FOLDER || this.plugin.isInboxFolder(this.folderSel.accountId, this.folderSel.folderId));
		const split = s.mailSplitInbox && triage;
		const sections = split ? splitSections(shown) : [{ key: "all" as SectionKey, label: "", messages: shown }];

		// conversations: group each section, then flatten back to the rows
		// that are actually visible. A conversation of one renders as a plain
		// message, so mail that was never a back-and-forth is untouched.
		// A search is never grouped: it asks about messages, and burying a
		// match inside a collapsed thread would misreport what matched.
		const threaded = s.mailConversations && !search;
		const collapsedSections = new Set(s.mailSectionsCollapsed);
		const plan: { header: MailSection | null; rows: MailRow[] }[] = [];
		const allThreads: MailThread[] = [];
		for (const sec of sections) {
			const rows: MailRow[] = [];
			if (threaded) {
				for (const t of groupThreads(sec.messages)) {
					allThreads.push(t);
					const many = t.messages.length > 1;
					const open = many && this.expanded.has(t.key);
					// a collapsed head answers for the whole thread; an expanded
					// one answers only for itself, since its children are their
					// own rows right underneath
					rows.push({ m: t.latest, targets: many && !open ? t.messages : [t.latest], thread: many ? t : null, child: false });
					if (open) for (const c of t.messages.slice(1)) rows.push({ m: c, targets: [c], thread: null, child: true });
				}
			} else for (const m of sec.messages) rows.push({ m, targets: [m], thread: null, child: false });
			plan.push({ header: split ? sec : null, rows: collapsedSections.has(sec.key) ? [] : rows });
		}
		this.lastThreads = allThreads;
		const rows = plan.flatMap((p) => p.rows);
		this.lastList = rows.map((r) => r.m);
		this.lastRows = rows.map((r) => r.targets);
		this.plugin.ensurePhotos(rows.map((r) => r.m.fromAddress));
		this.listEl.toggleClass("has-multisel", this.multiSel.size > 0);
		this.listEl.toggleClass("select-mode", this.selectMode);
		this.listEl.toggleClass("has-threads", threaded);
		for (const d of ["compact", "cozy", "comfortable"]) this.listEl.toggleClass(`density-${d}`, s.mailDensity === d);
		// the preview clamp rides a custom property so the rows themselves stay
		// identical whether one line is showing or three
		this.listEl.style.setProperty("--pcal-preview-lines", String(s.mailPreviewLines || 1));

		// Nothing below here is free. A row is about twenty elements and a dozen
		// listeners, and notify() fires on every poll, every body that lands and
		// every folder that syncs behind you. Rebuilding the list each time was
		// survivable at fifty messages and is not at fifteen hundred: marking one
		// message read would throw away and remake every other row on screen, and
		// take the scroll position with it.
		//
		// So each thing that can appear in the column says what it would say, and
		// only the ones whose answer changed are built. Mark one message read and
		// exactly one row is remade; the other fourteen hundred and ninety-nine
		// are left where they are.
		const view = [
			todayKey,
			emptyText ?? "",
			`${showTags ? 1 : 0}${threaded ? 1 : 0}${split ? 1 : 0}${this.selectMode ? 1 : 0}${s.use24h ? 1 : 0}${s.mailPhotos ? 1 : 0}`,
			`${s.mailDensity}|${s.mailPreviewLines}|${this.multiSel.size}`,
		].join(SIG_FIELD);
		const items: ListItem[] = [];
		errors.forEach((err, i) => items.push({ key: `err:${i}`, sig: err, make: () => this.listEl.createDiv({ cls: "pcal-mail-error", text: err }) }));
		if (emptyText) items.push({ key: "empty", sig: emptyText, make: () => this.listEl.createDiv({ cls: "pcal-embed-empty", text: emptyText }) });
		for (const part of plan) {
			if (part.header) {
				const sec = part.header;
				const secCollapsed = collapsedSections.has(sec.key);
				items.push({
					key: `sec:${sec.key}`,
					sig: `${sec.label}|${sec.messages.length}|${sec.messages.filter((x) => x.unread).length}|${secCollapsed ? 1 : 0}`,
					make: () => this.renderSectionHeader(sec, secCollapsed),
				});
			}
			part.rows.forEach((r, i) => {
				// a picked row whose neighbour below is also picked joins it into
				// one block, so that neighbour is part of what this row draws
				const below = part.rows[i + 1];
				const runNext = !!below && below.targets.every((x) => this.multiSel.has(x.id));
				items.push({
					key: `${r.child ? "c" : "r"}:${r.m.id}`,
					sig: `${view}${SIG_FIELD}${this.rowSignature(r, nameMaps)}${SIG_FIELD}${runNext ? 1 : 0}`,
					make: () => this.renderRow(r, todayKey, nameMaps, runNext),
				});
			});
		}
		this.reconcileList(items);
		// which row is open is deliberately absent from every signature above: it
		// is two class toggles, and signing it would remake rows for a change
		// paintSelection() already handles without touching one
		this.syncSelectedRowClasses();
		this.renderReading();
	}

	/** Bring the drawn column into line with `items`, building only what is new
	 *  or has changed and moving the rest into place.
	 *
	 *  Keyed by message id, so mail arriving at the top inserts rows rather than
	 *  shifting every signature down one and remaking the lot. */
	private reconcileList(items: ListItem[]) {
		this.drawnItems = reconcileChildren<Element>(this.listEl, items, this.drawnItems);
	}

	/** Everything one drawn row depends on, as one string.
	 *
	 *  It has to name every value renderRow reads, because a field left out is
	 *  a row that quietly stops updating. Which row is open is the deliberate
	 *  exception, handled by syncSelectedRowClasses instead. */
	private rowSignature(r: MailRow, nameMaps: Map<string, Map<string, string>> | null): string {
		const s = this.plugin.settings;
		const m = r.m;
		const t = r.thread;
		const collapsed = !!t && !this.expanded.has(t.key);
		const mk = collapsed && t ? t : m;
		return [
			m.id,
			collapsed && t ? (t.unread > 0 ? 1 : 0) : m.unread ? 1 : 0,
			r.child ? 1 : 0,
			r.targets.every((x) => this.multiSel.has(x.id)) ? 1 : 0,
			t ? `${t.key}/${t.messages.length}/${collapsed ? 1 : 0}/${t.senders.join("~")}` : "",
			m.from,
			m.subject,
			m.receivedMs,
			s.mailPreviewLines > 0 ? m.preview : "",
			`${mk.hasAttachments ? 1 : 0}${mk.flagged ? 1 : 0}${mk.priority ? 1 : 0}`,
			r.targets.every((x) => x.flagged) ? 1 : 0,
			(m.categories ?? []).map((c) => `${c}=${this.plugin.categoryColorFor(m.accountId, c)}`).join("~"),
			s.mailPhotos && this.plugin.photoFor(m.fromAddress) ? 1 : 0,
			(nameMaps && m.folderId ? nameMaps.get(m.accountId)?.get(m.folderId) : "") ?? "",
			// the actions the row draws depend on what it answers for
			r.targets.length,
		].join(SIG_FIELD);
	}

	/** A standing reminder while an out-of-office is answering for you.
	 *
	 *  This is the entire reason to build the feature here rather than leave
	 *  it to Outlook: an auto-reply nobody can see is one that runs for a
	 *  fortnight after you got back, telling everyone you are away. */
	private renderOooBar() {
		const bar = this.oooEl;
		if (!bar) return;
		bar.empty();
		const on = this.plugin.autoReplyingAccounts();
		if (!on.length) {
			bar.hide();
			return;
		}
		bar.show();
		setIcon(bar.createSpan("pcal-ooo-bar-icon"), "megaphone");
		bar.createSpan({
			cls: "pcal-ooo-bar-text",
			text: on.length === 1 ? `Automatic replies are on for ${this.plugin.nameOf(on[0])}.` : `Automatic replies are on for ${on.map((a) => this.plugin.nameOf(a)).join(" and ")}.`,
		});
		bar.createEl("button", { text: "Change" }).addEventListener("click", () => new OutOfOfficeModal(this.app, this.plugin).open());
		bar.createEl("button", { text: "Turn off" }).addEventListener("click", () => {
			for (const a of on) {
				const cur = this.plugin.autoReplyFor(a.id);
				if (cur) void this.plugin.saveAutoReply(a.id, { ...cur, status: "disabled" });
			}
		});
	}

	/** A section band in the split inbox: what it is, how much is in it, how
	 *  much of that is unread, and a twisty to fold it away. */
	private renderSectionHeader(sec: MailSection, collapsed: boolean): HTMLElement {
		const head = this.listEl.createDiv("pcal-mail-section");
		head.toggleClass("is-collapsed", collapsed);
		const tw = head.createSpan("pcal-mail-section-twist");
		setIcon(tw, collapsed ? "chevron-right" : "chevron-down");
		head.createSpan({ cls: "pcal-mail-section-name", text: sec.label });
		const unread = sec.messages.filter((m) => m.unread).length;
		head.createSpan({ cls: "pcal-mail-section-count", text: unread ? `${unread} of ${sec.messages.length}` : String(sec.messages.length) });
		head.setAttribute("aria-label", `${sec.label}, ${sec.messages.length} messages${unread ? `, ${unread} unread` : ""}`);
		head.addEventListener("click", () => {
			const set = new Set(this.plugin.settings.mailSectionsCollapsed);
			if (collapsed) set.delete(sec.key);
			else set.add(sec.key);
			this.plugin.settings.mailSectionsCollapsed = [...set];
			this.plugin.queueSave();
			this.render();
		});
		return head;
	}

	/** One row of the message list.
	 *
	 *  `thread` is set when the row heads a conversation of more than one,
	 *  which earns it the twisty, the message count, and marks rolled up off
	 *  every message in it. `targets` is what the checkbox and the quick
	 *  actions act on: a collapsed conversation answers for all of itself,
	 *  which is what "archive this thread" has to mean, and every other row
	 *  answers for itself alone. */
	private renderRow(r: MailRow, todayKey: string, nameMaps: Map<string, Map<string, string>> | null, runNext: boolean): HTMLElement {
		const s = this.plugin.settings;
		const m = r.m;
		const t = r.thread;
		const open = !!t && this.expanded.has(t.key);
		const collapsed = !!t && !open;
		const anyUnread = collapsed && t ? t.unread > 0 : m.unread;
		{
			const row = this.listEl.createDiv("pcal-mail-row");
			// whether the picked row below this one continues the block, so a run
			// of them drops its inner separators. Decided here, where the row's
			// neighbour is already known, rather than by a pass over the whole
			// list afterwards that a partial update would have no reason to run.
			row.toggleClass("is-multisel-run", runNext);
			// resting on a row is as good a signal as a click that this is the
			// one being read next, and it buys the fetch the time it takes to
			// travel from the row to the mouse button
			row.addEventListener("mouseenter", () => this.hoverWarm(m));
			row.addEventListener("mouseleave", () => this.clearHoverWarm());
			row.toggleClass("is-unread", anyUnread);
			row.toggleClass("is-selected", this.selected?.id === m.id);
			row.toggleClass("is-multisel", r.targets.every((x) => this.multiSel.has(x.id)));
			row.toggleClass("is-threadchild", r.child);
			// Outlook's read/unread gutter: the bar down the left edge is the
			// unread mark AND the control that clears it. It thickens under the
			// pointer so there is something to aim at, and a read row shows a
			// hollow bar on hover so marking one back unread is the same gesture
			// in the same place. Absolutely positioned, so widening it is paint
			// alone and the row never reflows as the mouse crosses the list.
			const gutter = row.createDiv("pcal-mail-readbar");
			gutter.createDiv("pcal-mail-readbar-fill");
			gutter.setAttribute("aria-label", anyUnread ? (collapsed ? "Mark conversation read" : "Mark read") : collapsed ? "Mark conversation unread" : "Mark unread");
			gutter.addEventListener("click", (e) => {
				e.stopPropagation();
				for (const x of r.targets) void this.plugin.setMailRead(x, anyUnread);
			});
			// the twisty column exists on every row so the avatars line up
			// whether or not a given row heads a conversation; CSS drops it
			// entirely when the list is not grouped at all
			const tw = row.createDiv("pcal-mail-twisty");
			if (t) {
				setIcon(tw, open ? "chevron-down" : "chevron-right");
				tw.setAttribute("aria-label", open ? "Collapse conversation" : `Expand conversation, ${t.messages.length} messages`);
				tw.addEventListener("click", (e) => {
					e.stopPropagation();
					if (open) this.expanded.delete(t.key);
					else this.expanded.add(t.key);
					this.render();
				});
			}
			// the avatar and the checkbox share a box, so the checkbox is
			// centered on the avatar by construction rather than by an offset
			// that anything added earlier in the row would silently break
			const avwrap = row.createDiv("pcal-mail-avatarwrap");
			const av = avwrap.createDiv("pcal-mail-avatar");
			// a real face when the directory has one, initials otherwise: the
			// colored circle is a stand-in, not the goal
			const photo = s.mailPhotos ? this.plugin.photoFor(m.fromAddress) : null;
			if (photo) {
				av.addClass("has-photo");
				av.style.backgroundImage = `url("${photo}")`;
			} else {
				av.style.backgroundColor = avatarColor(m.from || m.fromAddress);
				av.setText(avatarInitials(m.from || m.fromAddress));
			}
			av.setAttribute("aria-label", m.accountLabel);
			// the Outlook checkbox: over the avatar on hover, or always while a
			// selection is active; clicking it never changes the reading pane
			const check = avwrap.createDiv("pcal-mail-check");
			const allSel = r.targets.every((x) => this.multiSel.has(x.id));
			setIcon(check, allSel ? "check-square" : "square");
			check.addEventListener("click", (e) => {
				e.stopPropagation();
				this.setSelected(r.targets, !allSel);
				this.selAnchorId = m.id;
				this.render();
			});
			const mid = row.createDiv("pcal-mail-mid");
			const top = mid.createDiv("pcal-mail-top");
			// collapsed, the name line is everyone who wrote, newest first;
			// open, the head is just the newest message and its own sender
			top.createSpan({ cls: "pcal-mail-from", text: collapsed && t && t.senders.length > 1 ? t.senders.join(", ") : m.from });
			if (t) top.createSpan({ cls: "pcal-mail-count", text: String(t.messages.length), attr: { "aria-label": `${t.messages.length} messages in this conversation` } });
			const srcFolder = nameMaps && m.folderId ? nameMaps.get(m.accountId)?.get(m.folderId) : null;
			if (srcFolder) top.createSpan({ cls: "pcal-mail-tag", text: srcFolder });
			// Outlook's arrangement: the sender line carries the marks, the
			// subject line carries the date. The quick actions land on the
			// sender line on hover, so the marks are what steps aside for them
			// and the date stays readable the whole time.
			// a collapsed conversation wears the marks of everything inside it,
			// so an attachment three replies down is still visible from here
			const mk = collapsed && t ? t : m;
			const marks = top.createDiv("pcal-mail-marks");
			if (mk.hasAttachments) setIcon(marks.createSpan({ cls: "pcal-mail-clip", attr: { "aria-label": "Has attachments" } }), "paperclip");
			if (mk.flagged) {
				// the flag is its own target, the way Outlook's flag column
				// is: clicking it clears it without opening anything
				const fl = marks.createSpan({ cls: "pcal-mail-flagmark", attr: { "aria-label": "Flagged, click to clear" } });
				setIcon(fl, "flag");
				fl.addEventListener("click", (e) => {
					e.stopPropagation();
					for (const x of r.targets) if (x.flagged) void this.plugin.setMailFlag(x, false);
					this.render();
				});
			}
			if (mk.priority) marks.createSpan({ cls: "pcal-mail-bang", text: "!", attr: { "aria-label": "High importance" } });
			// categories as colored squares at the end of the marks, which is
			// where Outlook puts them and where the eye already looks
			for (const cat of m.categories ?? []) {
				const dot = marks.createSpan({ cls: "pcal-mail-cat", attr: { "aria-label": cat } });
				dot.style.backgroundColor = this.plugin.categoryColorFor(m.accountId, cat);
			}
			const subj = mid.createDiv("pcal-mail-subject");
			subj.createSpan({ cls: "pcal-mail-subject-text", text: m.subject });
			subj.createSpan({ cls: "pcal-mail-time", text: fmtMailTime(m.receivedMs, todayKey, s.use24h) });
			if (s.mailPreviewLines > 0) mid.createDiv({ cls: "pcal-mail-preview", text: m.preview });
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
			const flagOn = !r.targets.every((x) => x.flagged);
			act("flag", flagOn ? (collapsed ? "Flag conversation" : "Flag") : collapsed ? "Clear the conversation's flag" : "Clear flag", () => {
				for (const x of r.targets) void this.plugin.setMailFlag(x, flagOn);
				this.render();
			});
			act(anyUnread ? "mail-open" : "mail", collapsed ? (anyUnread ? "Mark conversation read" : "Mark conversation unread") : anyUnread ? "Mark read" : "Mark unread", () => {
				for (const x of r.targets) void this.plugin.setMailRead(x, anyUnread);
			});
			act("trash-2", collapsed ? "Delete conversation" : "Delete", () => this.deleteMessages(r.targets, false));
			row.addEventListener("click", (e) => {
				// Ctrl toggles membership, Shift takes the range from the last
				// anchor; a plain click clears the bulk set and reads normally
				if (e.ctrlKey || e.metaKey) {
					this.setSelected(r.targets, !r.targets.every((x) => this.multiSel.has(x.id)));
					this.selAnchorId = m.id;
					this.render();
					return;
				}
				if (e.shiftKey) {
					const anchor = this.selAnchorId ?? this.selected?.id ?? null;
					const ai = anchor ? this.lastList.findIndex((x) => x.id === anchor) : -1;
					const bi = this.lastList.findIndex((x) => x.id === m.id);
					if (ai >= 0 && bi >= 0) {
						// a range that crosses a collapsed conversation takes
						// the whole conversation, not just its newest message
						this.multiSel = new Set(
							this.lastRows
								.slice(Math.min(ai, bi), Math.max(ai, bi) + 1)
								.flat()
								.map((x) => x.id)
						);
						this.render();
					}
					return;
				}
				if (this.selectMode) {
					// in select mode the whole row is a checkbox
					this.setSelected(r.targets, !r.targets.every((x) => this.multiSel.has(x.id)));
					this.selAnchorId = m.id;
					this.render();
					return;
				}
				this.multiSel.clear();
				this.selAnchorId = m.id;
				void this.select(m);
			});
			// dragging a row onto a folder files it. A row already inside the
			// bulk selection drags the whole selection, which is what makes
			// tick-tick-tick-drag work.
			row.draggable = true;
			row.addEventListener("dragstart", (e) => {
				const dragging = this.multiSel.has(m.id) && this.multiSel.size ? this.multiTargets() : r.targets;
				this.mailDrag = dragging;
				e.dataTransfer?.setData("text/plain", dragging.map((x) => x.subject).join(", "));
				if (e.dataTransfer) e.dataTransfer.effectAllowed = "move";
			});
			row.addEventListener("dragend", () => (this.mailDrag = null));
			row.addEventListener("contextmenu", (e) => {
				e.preventDefault();
				// the menu acts on the whole selection when this row is in it,
				// and otherwise on whatever the row itself answers for
				const targets = this.multiSel.has(m.id) && this.multiSel.size ? this.multiTargets() : r.targets;
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
				const flagOn = !targets.every((x) => x.flagged);
				menu.addItem((i) =>
					i.setTitle(flagOn ? `Flag${plural}` : `Clear the flag${plural ? ` on${plural}` : ""}`).onClick(() => {
						for (const t of targets) void this.plugin.setMailFlag(t, flagOn);
						done(false);
					})
				);
				menu.addItem((i) => i.setTitle(`Categorize${plural}...`).onClick(() => this.categorizeMenu(targets)));
				if (targets.length === 1 && Platform.isDesktopApp) menu.addItem((i) => i.setTitle("Print...").setIcon("printer").onClick(() => void this.printMessage(targets[0])));
				menu.addItem((i) => i.setTitle(`Snooze${plural}`).onClick(() => this.snoozeTargets(targets)));
				menu.addItem((i) => i.setTitle(`Move${plural} to...`).onClick(() => this.openMovePicker(targets)));
				// a rule made from a message in front of you is how rules
				// actually get written, so the sender is filled in already
				menu.addItem((i) =>
					i.setTitle("Create a rule from this...").onClick(() => {
						new RuleEditModal(
							this.app,
							this.plugin,
							m.accountId,
							null,
							{ name: `From ${m.from}`, fromContains: m.fromAddress || m.from },
							() => this.render()
						).open();
					})
				);
				// the folders filed into lately, right on the menu, since the
				// same three or four take almost everything
				for (const r of this.plugin.settings.mailRecentFolders.slice(0, 4)) {
					if (!targets.some((t) => t.accountId === r.accountId)) continue;
					menu.addItem((i) =>
						i.setTitle(`Move${plural} to ${r.name}`).onClick(() => {
							void this.plugin.moveMail(targets, r.accountId, r.folderId, r.name).then(() => done(true));
						})
					);
				}
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
			return row;
		}
	}

	/** Add or drop a whole set of messages from the bulk selection at once,
	 *  which is what a collapsed conversation's checkbox needs. */
	private setSelected(targets: PCMail[], on: boolean) {
		for (const x of targets) {
			if (on) this.multiSel.add(x.id);
			else this.multiSel.delete(x.id);
		}
	}

	/** What the toolbar acts on: the bulk selection when one exists, else the
	 *  message in the reading pane. Resolved against every message the rows
	 *  answer for, not just the visible ones, so a collapsed conversation's
	 *  hidden replies are still in reach. */
	private multiTargets(): PCMail[] {
		if (this.multiSel.size) {
			const byId = new Map(this.lastRows.flat().map((x) => [x.id, x]));
			return [...this.multiSel].map((id) => byId.get(id)).filter((x): x is PCMail => !!x);
		}
		return this.selected ? [this.selected] : [];
	}

	/** The folder pane: each account as a group, its tree beneath, unread
	 *  counts riding along; All inboxes on top returns to the unified list. */
	private renderFolders(accounts: GraphAccount[], colorOf: Map<string, string>) {
		const host = this.foldersEl;
		if (!host) return;
		// the same bargain as the message list: this tree is rebuilt from
		// scratch on every notify, and a mail arriving in a folder you are not
		// looking at changes nothing here except a count that may not have moved
		const sig = this.folderSignature(accounts, colorOf);
		if (sig === this.lastFolderSig) return;
		this.lastFolderSig = sig;
		const scroll = host.scrollTop;
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
				const rn = folderIdCategory(f.folderId) ?? (f.folderId === UNREAD_FOLDER ? "Unread Mail" : this.plugin.folderNamesFor(ac.id).get(f.folderId) ?? "...");
				const dn = f.name?.trim() || rn;
				dispCounts.set(dn, (dispCounts.get(dn) ?? 0) + 1);
			}
			favs.forEach((fav, idx) => {
				const acc = accounts.find((x) => x.id === fav.accountId);
				if (!acc) return;
				const cat = folderIdCategory(fav.folderId);
				const isUnread = fav.folderId === UNREAD_FOLDER;
				const realName = cat ?? (isUnread ? "Unread Mail" : this.plugin.folderNamesFor(acc.id).get(fav.folderId) ?? "...");
				const name = fav.name?.trim() || realName;
				const count = cat ? this.plugin.categoryCount(acc.id, cat) : isUnread ? this.plugin.unreadSubtreeCount(acc) : this.plugin.folderUnreadRollup(acc.id, fav.folderId);
				const row = host.createDiv("pcal-folder-row pcal-fav-row");
				row.toggleClass("pcal-fav-indent", !!fav.indent);
				row.toggleClass("is-selected", this.folderSel?.accountId === acc.id && this.folderSel?.folderId === fav.folderId);
				// favorites are the folders filed into most, so they take a
				// dropped message too. A category takes one as well, and labels
				// it rather than moving it, which is what dropping onto a
				// category does in Outlook. Unread Mail is a search, not a place.
				if (cat) this.acceptCategoryDrop(row, acc.id, cat);
				else if (!isUnread) this.acceptMailDrop(row, acc.id, fav.folderId, name);
				if (cat) {
					const dot = row.createSpan({ cls: "pcal-folder-catdot" });
					dot.style.backgroundColor = this.plugin.categoryColorFor(acc.id, cat);
				} else {
					const ic = row.createSpan("pcal-folder-ic");
					setIcon(ic, isUnread ? "mail-open" : fav.folderId === this.plugin.inboxIdFor(acc) ? "inbox" : "folder");
				}
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
					menu.addItem((i) => i.setTitle("Remove from favorites").onClick(() => this.toggleFavorite(acc.id, fav.folderId)));
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
			// favorites are a shortlist over the top of the real mailboxes, so
			// they get a rule under them rather than running straight into the
			// first account's tree
			host.createDiv("pcal-folder-sep");
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
			head.addEventListener("contextmenu", (e) => {
				e.preventDefault();
				const menu = new Menu();
				menu.addItem((i) => i.setTitle("New folder...").setIcon("folder-plus").onClick(() => this.askNewFolder(a.id, null, this.plugin.nameOf(a))));
				menu.addSeparator();
				const arranged = this.plugin.folderOrderFor(a.id).length > 0;
				menu.addItem((i) =>
					i
						.setTitle("Sort folders alphabetically again")
						.setDisabled(!arranged)
						.onClick(() => this.plugin.resetFolderOrder(a.id))
				);
				menu.showAtMouseEvent(e);
			});
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
				this.acceptMailDrop(row, a.id, folder.id, folder.name);
				// folders rearrange by dragging one onto a sibling, which is
				// a display order only and never touches the mailbox
				row.draggable = true;
				row.addEventListener("dragstart", (e) => {
					this.folderDrag = { accountId: a.id, folderId: folder.id };
					e.stopPropagation();
					if (e.dataTransfer) e.dataTransfer.effectAllowed = "move";
				});
				row.addEventListener("dragend", () => {
					this.folderDrag = null;
					row.removeClass("pcal-drop-target");
				});
				row.addEventListener("dragover", (e) => {
					const d = this.folderDrag;
					if (!d || d.accountId !== a.id || d.folderId === folder.id) return;
					e.preventDefault();
					row.addClass("pcal-drop-target");
				});
				row.addEventListener("dragleave", () => row.removeClass("pcal-drop-target"));
				row.addEventListener("drop", (e) => {
					const d = this.folderDrag;
					row.removeClass("pcal-drop-target");
					if (!d || d.accountId !== a.id) return;
					e.preventDefault();
					e.stopPropagation();
					this.folderDrag = null;
					this.plugin.reorderFolder(a.id, d.folderId, folder.id);
				});
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
					menu.addItem((i) => i.setTitle(this.isFavorite(a.id, folder.id) ? "Remove from favorites" : "Add to favorites").onClick(() => this.toggleFavorite(a.id, folder.id)));
					menu.addItem((i) =>
						i.setTitle("Hide folder").onClick(() => {
							this.plugin.settings.mailHiddenFolders = [...this.plugin.settings.mailHiddenFolders, { accountId: a.id, folderId: folder.id }];
							this.plugin.queueSave();
							this.render();
						})
					);
					menu.addSeparator();
					menu.addItem((i) => i.setTitle("New subfolder...").setIcon("folder-plus").onClick(() => this.askNewFolder(a.id, folder.id, folder.name)));
					// the mailbox runs on its own folders; renaming Sent Items
					// or deleting Drafts is not on offer
					const system = isSystemFolder(folder.name, folder.id, inboxId);
					menu.addItem((i) =>
						i
							.setTitle("Rename...")
							.setIcon("pencil")
							.setDisabled(system)
							.onClick(() => {
								new PromptModal(this.app, `Rename ${folder.name}`, [{ label: "Name", value: folder.name }], ([name]) => {
									if (name.trim() && name.trim() !== folder.name) void this.plugin.renameFolder(a.id, folder.id, name);
								}).open();
							})
					);
					menu.addItem((i) =>
						i
							.setTitle("Delete folder")
							.setIcon("trash-2")
							.setDisabled(system)
							.onClick(() => this.askDeleteFolder(a.id, folder.id, folder.name))
					);
					menu.showAtMouseEvent(e);
				});
			}

			// hidden folders come back through their own quiet row
			const hidden = this.plugin.settings.mailHiddenFolders.filter((h) => h.accountId === a.id);
			if (hidden.length) {
				const hr = host.createDiv("pcal-folder-row pcal-folder-hiddenrow");
				hr.addClass("pcal-depth-0");
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

			// Categories: the mailbox's own list, each one a place to open
			// rather than only a label to apply.
			//
			// The branch is drawn even when the list is empty, which it was not
			// at first. Hiding it meant a mailbox with no categories yet showed
			// nothing at all, and the way to make one lives behind this very
			// branch, so the feature was invisible to exactly the people who had
			// not used it. An empty branch says which of the two reasons it is
			// empty for, and offers the way out of both.
			const cats = this.plugin.categoriesFor(a.id);
			{
				const catKey = `cats:${a.id}`;
				const catCollapsed = collapsed.has(catKey);
				const chead = host.createDiv("pcal-folder-row pcal-folder-catroot");
				chead.addClass("pcal-depth-0");
				const ctw = chead.createSpan("pcal-folder-twist");
				setIcon(ctw, catCollapsed ? "chevron-right" : "chevron-down");
				const cic = chead.createSpan("pcal-folder-ic");
				setIcon(cic, "tag");
				chead.createSpan({ cls: "pcal-folder-name", text: "Categories" });
				chead.addEventListener("click", () => toggleCollapse(catKey));
				chead.addEventListener("contextmenu", (e) => {
					e.preventDefault();
					const menu = new Menu();
					menu.addItem((i) =>
						i
							.setTitle("New category...")
							.setIcon("plus")
							.setDisabled(!this.plugin.canReadCategories(a.id))
							.onClick(() => this.askNewCategory(a.id))
					);
					menu.addItem((i) => i.setTitle("Manage categories...").setIcon("settings-2").onClick(() => new CategoriesModal(this.app, this.plugin, () => this.render()).open()));
					menu.showAtMouseEvent(e);
				});
				if (!catCollapsed) {
					for (const cat of cats) {
						const fid = categoryFolderId(cat.displayName);
						const row = host.createDiv("pcal-folder-row");
						row.addClass("pcal-depth-1");
						row.toggleClass("is-selected", this.folderSel?.accountId === a.id && this.folderSel?.folderId === fid);
						row.createSpan("pcal-folder-twist"); // the empty chevron slot every tree row keeps
						const dot = row.createSpan({ cls: "pcal-folder-catdot" });
						dot.style.backgroundColor = categoryColor(cat.color);
						row.createSpan({ cls: "pcal-folder-name", text: cat.displayName });
						const held = this.plugin.categoryCount(a.id, cat.displayName);
						if (held > 0) row.createSpan({ cls: "pcal-folder-count", text: String(held) });
						this.acceptCategoryDrop(row, a.id, cat.displayName);
						row.addEventListener("click", () => {
							this.plugin.clearMailSearch();
							this.folderSel = { accountId: a.id, folderId: fid, name: cat.displayName };
							this.screen = "list";
							this.autoSelectPending = true;
							this.plugin.ensureFolderMail(a.id, fid, false);
							this.render();
						});
						row.addEventListener("contextmenu", (e) => {
							e.preventDefault();
							const menu = new Menu();
							menu.addItem((i) => i.setTitle(this.isFavorite(a.id, fid) ? "Remove from favorites" : "Add to favorites").onClick(() => this.toggleFavorite(a.id, fid)));
							menu.addItem((i) => i.setTitle("Manage categories...").setIcon("settings-2").onClick(() => new CategoriesModal(this.app, this.plugin, () => this.render()).open()));
							menu.showAtMouseEvent(e);
						});
					}
					// an account that cannot read its categories is not the same
					// as one with none, and saying "New category" to the first
					// would offer something the mailbox is going to refuse
					if (!cats.length && !this.plugin.canReadCategories(a.id)) {
						const note = host.createDiv("pcal-folder-row pcal-folder-addcat");
						note.addClass("pcal-depth-1");
						note.createSpan("pcal-folder-twist");
						const ic = note.createSpan("pcal-folder-ic");
						setIcon(ic, "key-round");
						note.createSpan({ cls: "pcal-folder-name", text: "Reconnect to read categories" });
						note.addEventListener("click", () => this.plugin.openOwnSettings());
					} else if (!cats.length) {
						const add = host.createDiv("pcal-folder-row pcal-folder-addcat");
						add.addClass("pcal-depth-1");
						add.createSpan("pcal-folder-twist");
						const ic = add.createSpan("pcal-folder-ic");
						setIcon(ic, "plus");
						add.createSpan({ cls: "pcal-folder-name", text: "New category" });
						add.addEventListener("click", () => this.askNewCategory(a.id));
					}
				}
			}

			// Search Folders close out the account, Outlook-style: a collapsible
			// branch holding the virtual Unread view and any saved searches
			const searchKey = `search:${a.id}`;
			const searchCollapsed = collapsed.has(searchKey);
			const shead = host.createDiv("pcal-folder-row pcal-folder-searchroot");
			shead.addClass("pcal-depth-0");
			const stw = shead.createSpan("pcal-folder-twist");
			setIcon(stw, searchCollapsed ? "chevron-right" : "chevron-down");
			const sic = shead.createSpan("pcal-folder-ic");
			setIcon(sic, "search");
			shead.createSpan({ cls: "pcal-folder-name", text: "Search Folders" });
			shead.addEventListener("click", () => toggleCollapse(searchKey));
			if (searchCollapsed) continue;
			const unreadRow = host.createDiv("pcal-folder-row pcal-folder-unread");
			unreadRow.addClass("pcal-depth-1");
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
				menu.addItem((i) => i.setTitle(this.isFavorite(a.id, UNREAD_FOLDER) ? "Remove from favorites" : "Add to favorites").onClick(() => this.toggleFavorite(a.id, UNREAD_FOLDER)));
				menu.showAtMouseEvent(e);
			});
			for (const sf of this.plugin.settings.mailSearchFolders.filter((x) => x.accountId === a.id)) {
				const row = host.createDiv("pcal-folder-row");
				row.addClass("pcal-depth-1");
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
			addSearch.addClass("pcal-depth-1");
			addSearch.createSpan("pcal-folder-twist");
			const asIc = addSearch.createSpan("pcal-folder-ic");
			setIcon(asIc, "plus");
			addSearch.createSpan({ cls: "pcal-folder-name", text: "New search folder" });
			addSearch.addEventListener("click", () => new SearchFolderModal(this.app, this.plugin, a, null, () => this.render()).open());
		}
		if (scroll > 0) host.scrollTop = scroll;
	}

	/** Everything the folder pane is drawn from, as one string: which folders
	 *  exist, what they are called, what they are holding unread, which of them
	 *  are folded away, pinned, hidden or open. */
	private folderSignature(accounts: GraphAccount[], colorOf: Map<string, string>): string {
		const s = this.plugin.settings;
		const collapsed = new Set(s.mailCollapsed);
		const parts: string[] = [
			this.folderSel ? `${this.folderSel.accountId}/${this.folderSel.folderId}/${this.folderSel.name}` : "-",
			s.mailCollapsed.join(","),
			s.mailHiddenFolders.map((h) => `${h.accountId}/${h.folderId}`).join(","),
			s.mailFavorites.map((f) => `${f.accountId}/${f.folderId}/${f.name ?? ""}/${f.indent ? 1 : 0}`).join(","),
			s.mailSearchFolders.map((f) => `${f.accountId}/${f.id}/${f.name}`).join(","),
		];
		for (const a of accounts) {
			parts.push(`@${a.id}|${this.plugin.nameOf(a)}|${colorOf.get(a.id) ?? ""}|${this.plugin.inboxIdFor(a) ?? ""}|${this.plugin.unreadSubtreeCount(a)}`);
			// categories land after the tree does, and a pinned one wears its
			// color up in Favorites, so this is read even for a folded account.
			// The scope rides along because an empty branch says something
			// different depending on it, and a reconnect changes it underneath.
			parts.push(
				`${this.plugin.canReadCategories(a.id) ? 1 : 0}|` +
					this.plugin.categoriesFor(a.id).map((c) => `${c.displayName}=${c.color}=${this.plugin.categoryCount(a.id, c.displayName)}`).join("~")
			);
			if (collapsed.has(`acct:${a.id}`)) continue;
			const tree = this.plugin.folderTreeFor(a);
			if (!tree.length) parts.push("(loading)");
			for (const { folder, depth, expandable } of tree) {
				// a folded branch shows its subtree's total, so that is what has
				// to be compared for it rather than its own count
				const count = collapsed.has(folder.id) && expandable ? this.plugin.folderUnreadRollup(a.id, folder.id) : folder.unread;
				parts.push(`${folder.id}|${folder.name}|${count}|${depth}|${expandable ? 1 : 0}`);
			}
		}
		return parts.join(SIG_ROW);
	}

	/** Let a folder row take a dragged message. Only the rows of the account
	 *  the messages belong to light up, since a move cannot cross mailboxes
	 *  and a target that would silently do nothing should not look willing. */
	private acceptMailDrop(row: HTMLElement, accountId: string, folderId: string, name: string) {
		const mine = () => (this.mailDrag ?? []).filter((m) => m.accountId === accountId);
		row.addEventListener("dragover", (e) => {
			if (!mine().length) return;
			e.preventDefault();
			if (e.dataTransfer) e.dataTransfer.dropEffect = "move";
			row.addClass("pcal-mail-drop");
		});
		row.addEventListener("dragleave", () => row.removeClass("pcal-mail-drop"));
		row.addEventListener("drop", (e) => {
			row.removeClass("pcal-mail-drop");
			const targets = mine();
			if (!targets.length) return;
			e.preventDefault();
			e.stopPropagation();
			this.mailDrag = null;
			void this.plugin.moveMail(targets, accountId, folderId, name).then((n) => {
				if (n) new Notice(n > 1 ? `Power Desk: moved ${n} messages to ${name}.` : `Power Desk: moved to ${name}.`);
				if (targets.some((t) => t.id === this.selected?.id)) this.selected = null;
				this.multiSel.clear();
				this.render();
			});
		});
	}

	/** Let a category row take a dragged message. Dropping labels the message
	 *  rather than moving it, so it stays filed where it was and the notice
	 *  says tagged rather than moved. Dropping something that already carries
	 *  the category is not a mistake and does not report as one: what was
	 *  asked for is the end state, and it is already true. */
	private acceptCategoryDrop(row: HTMLElement, accountId: string, category: string) {
		const mine = () => (this.mailDrag ?? []).filter((m) => m.accountId === accountId);
		row.addEventListener("dragover", (e) => {
			if (!mine().length) return;
			e.preventDefault();
			// copy, not move: the message is not going anywhere
			if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
			row.addClass("pcal-mail-drop");
		});
		row.addEventListener("dragleave", () => row.removeClass("pcal-mail-drop"));
		row.addEventListener("drop", (e) => {
			row.removeClass("pcal-mail-drop");
			const dragged = mine();
			if (!dragged.length) return;
			e.preventDefault();
			e.stopPropagation();
			this.mailDrag = null;
			const fresh = dragged.filter((m) => !inCategory(m, category));
			for (const t of fresh) void this.plugin.setMailCategories(t, toggleCategory(t.categories, category));
			if (fresh.length) new Notice(fresh.length > 1 ? `Power Desk: tagged ${fresh.length} messages ${category}.` : `Power Desk: tagged ${category}.`);
			this.multiSel.clear();
			this.render();
		});
	}

	/** Make a category from the tree, without a trip through the manager. The
	 *  color is not asked for: the mailbox's first unused one is a good answer
	 *  and the manager can recolor it later. */
	private askNewCategory(accountId: string) {
		new PromptModal(this.app, "New category", [{ label: "Name", value: "", placeholder: "Waiting on" }], ([name]) => {
			if (!name.trim()) return;
			void this.plugin.newCategory(accountId, name, this.plugin.freeCategoryColor(accountId)).then(() => this.render());
		}).open();
	}

	private askNewFolder(accountId: string, parentId: string | null, whereName: string) {
		new PromptModal(this.app, parentId ? `New folder in ${whereName}` : `New folder in ${whereName}`, [{ label: "Name", value: "", placeholder: "Clients" }], ([name]) => {
			if (name.trim()) void this.plugin.newFolder(accountId, name, parentId);
		}).open();
	}

	/** Deleting a folder takes its mail with it, so the question says how
	 *  much that is. What the mailbox does with the contents afterwards is
	 *  Exchange's business and not something to promise here, so the wording
	 *  points at Outlook's own recovery rather than guaranteeing one. */
	private askDeleteFolder(accountId: string, folderId: string, name: string) {
		const w = this.plugin.folderWeight(accountId, folderId);
		const parts: string[] = [];
		if (w.messages) parts.push(`${w.messages} message${w.messages === 1 ? "" : "s"}`);
		if (w.folders) parts.push(`${w.folders} subfolder${w.folders === 1 ? "" : "s"}`);
		const holding = parts.length ? `It holds ${parts.join(" and ")}, which go with it. ` : "It looks empty. ";
		new ConfirmModal(
			this.app,
			`Delete the ${name} folder?`,
			`${holding}If you need any of it back afterwards, Recover Deleted Items in Outlook is where to look.`,
			"Delete folder",
			() => {
				if (this.folderSel?.accountId === accountId && this.folderSel.folderId === folderId) {
					this.folderSel = null;
					this.selected = null;
				}
				void this.plugin.removeFolder(accountId, folderId).then(() => this.render());
			}
		).open();
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
		// a warm body goes in before the first paint, so a prefetched message
		// draws with the click rather than a frame after it
		this.selectedBody = this.plugin.cachedMailBody(m);
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
		// the phone's drill moves between screens, which is a real layout
		// change; on a desktop the list is unchanged and only two classes and
		// the reading pane move
		if (this.drill || !this.paintSelection()) this.render();
		// whatever you read next is almost always next in the list, so warm
		// the way you are travelling while this one is on screen
		this.warmNeighbors(m);
		const body = this.selectedBody ?? (await this.plugin.readMailBody(m));
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
			// fetch supplies it alongside the bytes. All at once rather than
			// one after another: a mail with a signature logo and four badges
			// used to draw them in five round trips end to end.
			const html = body?.html ?? "";
			const inline = atts.filter((a) => a.isInline).slice(0, 8);
			if (inline.length && html.includes("cid:")) {
				const raws = await Promise.all(inline.map((att) => this.plugin.mailAttachmentRaw(m, att.id, true)));
				if (this.selected?.id !== m.id) return;
				const map = new Map<string, string>();
				for (const raw of raws) {
					if (raw?.contentId && html.includes(raw.contentId)) map.set(raw.contentId, `data:${raw.contentType};base64,${raw.contentBytes}`);
				}
				if (map.size) {
					this.inlineCids = map;
					this.renderReading();
				}
			}
		}
	}

	/** Move the selection without rebuilding the list.
	 *
	 *  A click changes one class on two rows and the contents of the reading
	 *  pane. The full render also rebuilds the folder tree, every row, every
	 *  avatar and every badge, which on a real mailbox is thousands of nodes
	 *  thrown away and made again to say "this one now". False when the rows
	 *  on screen no longer line up with the list they were built from, which
	 *  is the one case that does need the real thing. */
	private paintSelection(): boolean {
		if (!this.syncSelectedRowClasses()) return false;
		this.renderReading();
		return true;
	}

	/** Move the is-selected class onto the open row and off every other one,
	 *  without touching anything else. False when the drawn rows no longer line
	 *  up with the list they were built from, which is the one case that needs
	 *  the real thing.
	 *
	 *  Which row is open is deliberately absent from the list signature, so a
	 *  render that follows a click does not rebuild the list just to move two
	 *  classes; this is what puts them where they belong instead. */
	private syncSelectedRowClasses(): boolean {
		const rows = this.listEl?.querySelectorAll(".pcal-mail-row");
		if (!rows || rows.length !== this.lastList.length) return false;
		for (let i = 0; i < rows.length; i++) (rows[i] as HTMLElement).toggleClass("is-selected", this.lastList[i].id === this.selected?.id);
		return true;
	}

	/** Both views carry this button, and one of them is always the tab you are
	 *  not looking at; a button that lies about the state of the thing it
	 *  toggles is worse than no button. */
	private syncFoldToggle() {
		this.foldToggleBtn?.toggleClass("is-active", this.plugin.focusOn());
	}

	/** Hovering warms the body, but only once the pointer settles: sweeping
	 *  the mouse down a list crosses thirty rows and none of them is being
	 *  read. One timer for the whole view, so there is nothing to leak. */
	private hoverTimer: number | null = null;

	private hoverWarm(m: PCMail) {
		this.clearHoverWarm();
		this.hoverTimer = window.setTimeout(() => {
			this.hoverTimer = null;
			this.plugin.warmBody(m);
		}, 140);
	}

	private clearHoverWarm() {
		if (this.hoverTimer != null) window.clearTimeout(this.hoverTimer);
		this.hoverTimer = null;
	}

	/** Warm what is about to be read: the next few messages down the list and
	 *  the one above, since reading is a direction and not a random walk. */
	private warmNeighbors(m: PCMail) {
		const i = this.lastList.findIndex((x) => x.id === m.id);
		if (i < 0) return;
		for (const j of [i + 1, i + 2, i + 3, i - 1]) {
			const n = this.lastList[j];
			if (n) this.plugin.warmBody(n);
		}
	}

	private async openAttachment(m: PCMail, att: MailAttachment, open: boolean) {
		const path = await this.plugin.saveMailAttachment(m, att);
		if (!path) return;
		new Notice(`Power Desk: saved ${path}.`);
		if (open) (this.app as unknown as { openWithDefaultApp?: (p: string) => void }).openWithDefaultApp?.(path);
	}

	/** Preview means inside Obsidian: PDFs, images, and text all render in a
	 *  tab, which is the whole point of reading mail here rather than in
	 *  Outlook. A type Obsidian cannot draw falls through to the OS, so the
	 *  menu item never simply does nothing. */
	/** Unsubscribing always asks first, and the question names the exact
	 *  destination.
	 *
	 *  These addresses and URLs come from whoever sent the mail, not from
	 *  anything this plugin knows, so none of them is acted on unseen: a
	 *  one-click POST reaches a stranger's server, an https link opens their
	 *  page, and a mailto sends mail from your mailbox with your address on
	 *  it. Each is worth a glance before it happens, and a link disguised as
	 *  an unsubscribe is exactly the sort of thing that shows up in mail. */
	private confirmUnsubscribe(m: PCMail, u: UnsubscribeInfo) {
		const plan = unsubscribePlan(u);
		const what =
			plan.kind === "post"
				? `Power Desk will tell ${plan.target} to stop sending. Nothing else about you is sent with it.`
				: plan.kind === "open"
					? `This opens ${plan.target} in your browser, where the sender finishes it. Check the address looks like it belongs to them.`
					: `This sends mail from your mailbox to ${plan.target}, which tells the sender your address is real and read.`;
		new ConfirmModal(
			this.app,
			`Unsubscribe from ${m.from}?`,
			`${what} This came from the message itself, so it is the sender's own claim about where to go rather than anything Power Desk can vouch for.`,
			"Unsubscribe",
			() => void this.doUnsubscribe(m, u)
		).open();
	}

	private async doUnsubscribe(m: PCMail, u: UnsubscribeInfo) {
		const plan = unsubscribePlan(u);
		if (plan.kind === "open") {
			window.open(plan.target, "_blank");
			return;
		}
		if (plan.kind === "post") {
			try {
				await postOneClickUnsubscribe(plan.target);
				new Notice(`Power Desk: asked ${m.from} to stop sending. It can take a few days to take effect.`);
			} catch (e) {
				new Notice(`Power Desk: that unsubscribe did not go through. ${e instanceof Error ? e.message : String(e)}${u.webUrl ? " Try the link in the message itself." : ""}`, 9000);
			}
			return;
		}
		const sender = this.plugin.mailSender(m.accountId);
		if (!sender || !u.mailto) {
			new Notice("Power Desk: no account can send the unsubscribe mail.");
			return;
		}
		try {
			await sender.send({ to: [u.mailto.to], subject: u.mailto.subject, html: u.mailto.body });
			new Notice(`Power Desk: unsubscribe request sent to ${u.mailto.to}.`);
		} catch (e) {
			new Notice("Power Desk: " + (e instanceof Error ? e.message : String(e)));
		}
	}

	private async previewAttachment(m: PCMail, att: MailAttachment) {
		const path = await this.plugin.saveMailAttachment(m, att);
		if (!path) return;
		const f = this.app.vault.getAbstractFileByPath(path);
		const viewable = /\.(pdf|png|jpe?g|gif|webp|bmp|avif|svg|md|txt|mp4|webm|mov|mp3|wav|m4a|ogg|flac)$/i.test(path);
		if (f instanceof TFile && viewable) {
			await this.app.workspace.getLeaf(true).openFile(f);
			return;
		}
		new Notice(`Power Desk: saved ${path}.`);
		(this.app as unknown as { openWithDefaultApp?: (p: string) => void }).openWithDefaultApp?.(path);
	}

	private async saveAllAttachments(m: PCMail, files: MailAttachment[]) {
		const saved: string[] = [];
		for (const att of files) {
			const path = await this.plugin.saveMailAttachment(m, att);
			if (path) saved.push(path);
		}
		if (!saved.length) new Notice("Power Desk: none of those attachments carried a file to save.");
		else new Notice(`Power Desk: saved ${saved.length} attachment${saved.length === 1 ? "" : "s"} to ${saved[0].replace(/\/[^/]*$/, "") || "the vault"}.`);
	}

	/** An attachment's actions, in Outlook's order, dropped directly under the
	 *  chip rather than wherever the pointer happened to be: the menu belongs
	 *  to the file, so it should hang off it like a select. */
	private attachmentMenu(chip: HTMLElement, m: PCMail, att: MailAttachment, files: MailAttachment[]) {
		const menu = new Menu();
		menu.addItem((i) => i.setTitle("Preview").setIcon("eye").onClick(() => void this.previewAttachment(m, att)));
		if (Platform.isDesktopApp) menu.addItem((i) => i.setTitle("Open").setIcon("external-link").onClick(() => void this.openAttachment(m, att, true)));
		menu.addSeparator();
		menu.addItem((i) => i.setTitle("Save to vault").setIcon("folder-input").onClick(() => void this.openAttachment(m, att, false)));
		if (Platform.isDesktopApp) menu.addItem((i) => i.setTitle("Save to folder").setIcon("hard-drive").onClick(() => void this.plugin.saveMailAttachmentLocal(m, att)));
		if (files.length > 1) menu.addItem((i) => i.setTitle(`Save all ${files.length} to vault`).setIcon("download").onClick(() => void this.saveAllAttachments(m, files)));
		menu.addSeparator();
		menu.addItem((i) =>
			i
				.setTitle("Copy name")
				.setIcon("copy")
				.onClick(() => void navigator.clipboard.writeText(att.name))
		);
		// sized and placed to hang off the chip like a select rather than
		// float beside it: same left edge, same width, and close enough that
		// the two read as one control
		const r = chip.getBoundingClientRect();
		const dom = (menu as unknown as { dom?: HTMLElement }).dom;
		if (dom) {
			dom.addClass("pcal-att-menu");
			dom.style.setProperty("--pcal-att-menu-w", `${Math.round(r.width)}px`);
		}
		menu.showAtPosition({ x: r.left, y: r.bottom + 1 });
	}

	private async selectAndReply(m: PCMail, all = false) {
		await this.select(m);
		new RichComposeModal(this.app, this.plugin, { mode: all ? "replyAll" : "reply", mail: m }).open();
	}

	private renderReading() {
		const host = this.readEl;
		if (!host) return;
		const m = this.selected;
		for (const b of this.mailToolBtns) b.toggleClass("is-disabled", !m && this.multiSel.size === 0);
		// Rebuilding this pane means sanitizing and re-parsing the whole message
		// body, which for a newsletter is a great deal of HTML, and it takes the
		// reading position with it: scroll halfway down a long mail, wait for the
		// poll, and you are back at the top. It is called at the end of every
		// render, so most of the time the answer is the same pane it already
		// drew. Body text for a given id never changes, so a short description
		// of what went in is enough to know nothing has to come out.
		const sig = this.readingSignature(m);
		if (sig === this.lastReadSig) return;
		this.lastReadSig = sig;
		host.empty();
		if (!m) {
			host.createDiv({ cls: "pcal-embed-empty", text: "Select a message." });
			return;
		}
		host.createDiv({ cls: "pcal-mail-read-subject", text: m.subject });
		const fromRow = host.createDiv({ cls: "pcal-mail-read-meta" });
		fromRow.createSpan({ text: `${m.from}${m.fromAddress ? ` <${m.fromAddress}>` : ""}` });
		// bulk senders are required to say how to leave; when this one did,
		// the way out sits beside their name rather than at the bottom of the
		// mail in six-point grey, which is the whole reason it is hard to find
		const unsub = this.selectedBody?.unsub;
		if (unsub) {
			const link = fromRow.createSpan({ cls: "pcal-unsub", text: "Unsubscribe" });
			link.setAttribute("aria-label", "Unsubscribe from this sender");
			link.addEventListener("click", () => this.confirmUnsubscribe(m, unsub));
		}
		if (this.selectedBody?.toLine) host.createDiv({ cls: "pcal-mail-read-meta", text: `To: ${this.selectedBody.toLine}` });
		// categories read as named pills here, where there is room for words
		if ((m.categories ?? []).length) {
			const bar = host.createDiv("pcal-mail-cats");
			for (const cat of m.categories ?? []) {
				const pill = bar.createSpan({ cls: "pcal-mail-cat-pill", text: cat });
				const color = this.plugin.categoryColorFor(m.accountId, cat);
				pill.style.setProperty("--pcal-cat-color", color);
			}
		}

		// the rest of the conversation, newest first, so a back-and-forth
		// reads without going back to the list for every reply
		const convo = this.lastThreads.find((t) => t.messages.length > 1 && t.messages.some((x) => x.id === m.id));
		if (convo) {
			const todayKey = keyOfDate(new Date());
			const strip = host.createDiv("pcal-mail-convo");
			strip.createDiv({ cls: "pcal-mail-convo-head", text: `Conversation, ${convo.messages.length} messages` });
			for (const x of convo.messages) {
				const it = strip.createDiv("pcal-mail-convo-item");
				it.toggleClass("is-current", x.id === m.id);
				it.toggleClass("is-unread", x.unread);
				it.createSpan({ cls: "pcal-mail-convo-from", text: x.from });
				it.createSpan({ cls: "pcal-mail-convo-time", text: fmtMailTime(x.receivedMs, todayKey, this.plugin.settings.use24h) });
				if (x.id !== m.id) it.addEventListener("click", () => void this.select(x));
			}
		}

		// no button row here: everything it carried lives on the toolbar at the
		// top, which acts on the same message and is customizable

		// attachments ride their own bar, Outlook-style
		const files = (this.selectedAtts ?? []).filter((a) => !a.isInline);
		if ((m.hasAttachments && !this.selectedAtts) || files.length) {
			const bar = host.createDiv("pcal-mail-atts");
			if (!this.selectedAtts) bar.createSpan({ cls: "pcal-mail-att-note", text: "Loading attachments..." });
			if (files.length > 1) {
				const all = bar.createDiv("pcal-mail-att-all");
				setIcon(all.createSpan("pcal-mail-att-all-icon"), "download");
				all.createSpan({ text: `Save all ${files.length}` });
				all.setAttribute("aria-label", "Save every attachment to the vault");
				all.addEventListener("click", () => void this.saveAllAttachments(m, files));
			}
			for (const att of files) {
				const chip = bar.createDiv("pcal-mail-att");
				const badge = attachmentBadge(att.name, att.contentType);
				const ic = chip.createSpan({ cls: "pcal-mail-att-icon", text: badge.label });
				ic.style.setProperty("--pcal-att-color", badge.color);
				const tx = chip.createDiv("pcal-mail-att-text");
				tx.createDiv({ cls: "pcal-mail-att-name", text: att.name });
				tx.createDiv({ cls: "pcal-mail-att-size", text: fmtAttachmentSize(att.size) });
				const chev = chip.createSpan({ cls: "pcal-mail-att-chevron", attr: { "aria-label": `Actions for ${att.name}` } });
				setIcon(chev, "chevron-down");
				chip.setAttribute("aria-label", `${att.name}, ${fmtAttachmentSize(att.size)}`);
				// the whole chip opens the menu, exactly as the chevron does:
				// Outlook makes you find the arrow, and there is no reason to
				chip.addEventListener("click", () => this.attachmentMenu(chip, m, att, files));
			}
		}


		// HTML mail renders sanitized on its own light card; plain text stays text
		let html = this.selectedBody?.html ?? "";
		if (html && this.inlineCids) for (const [cid, url] of this.inlineCids) html = html.split(`cid:${cid}`).join(url);
		const bodyHost = host.createDiv("pcal-mail-read-body");
		if (html) bodyHost.createDiv("pcal-mail-html").appendChild(sanitizeHTMLToDom(html));
		else if (this.selectedBody) bodyHost.setText(this.selectedBody.text);
		// the list already knows the first line or two, so a cold message shows
		// its own opening while the rest travels, rather than the word Loading
		else if (this.selected?.preview) bodyHost.createDiv({ cls: "pcal-mail-read-pending", text: this.selected.preview });
		else bodyHost.setText("Loading...");
	}

	/** Everything the reading pane is drawn from, as one string. Bodies are
	 *  immutable per message id, so their lengths stand in for their contents;
	 *  everything else that shows here is named outright. */
	private readingSignature(m: PCMail | null): string {
		if (!m) return "-";
		const b = this.selectedBody;
		const convo = this.lastThreads.find((t) => t.messages.length > 1 && t.messages.some((x) => x.id === m.id));
		return [
			m.id,
			m.subject,
			m.from,
			m.fromAddress,
			m.hasAttachments ? 1 : 0,
			(m.categories ?? []).map((c) => `${c}=${this.plugin.categoryColorFor(m.accountId, c)}`).join("~"),
			b ? `${b.html ? `h${b.html.length}` : `t${b.text.length}`}|${b.toLine}|${b.unsub ? 1 : 0}` : "-",
			this.selectedAtts ? this.selectedAtts.map((a) => `${a.id}:${a.isInline ? 1 : 0}`).join("~") : "-",
			this.inlineCids ? this.inlineCids.size : "-",
			convo ? convo.messages.map((x) => `${x.id}${x.unread ? "!" : ""}`).join("~") : "-",
			this.plugin.settings.use24h ? 1 : 0,
		].join(SIG_FIELD);
	}
}

/** One shape a thing can take on paper.
 *
 *  The builder is handed whatever the window is set to and returns a whole
 *  document, so a style that has no use for an option is free to ignore it.
 *  `basePt` is the size that style has always printed at, which is what the
 *  window calls Normal: nothing about the defaults moved when the options
 *  arrived. */
interface PrintStyle {
	id: string;
	label: string;
	hint?: string;
	icon: string;
	basePt: number;
	/** How this style wants the paper before anyone says otherwise. */
	landscape: boolean;
	build: (o: PrintOptions) => string;
}

/** The print window: pick a style, look at the page, print the page you saw.
 *
 *  Outlook asks before it prints rather than throwing the system dialog up at
 *  the first click, and it is right to. A message and a list of messages want
 *  different paper, a month wants it sideways, and finding any of that out
 *  from the printer costs a sheet each time. What is previewed here is the
 *  document itself in a frame, and Print prints that same frame, so there is
 *  no second rendering path that can disagree with the preview. */
class PrintModal extends Modal {
	private styleId: string;
	private scale: PrintScaleId = "m";
	/** Per style, because portrait is right for a message and wrong for a
	 *  month grid; unset means the style's own preference. */
	private orient = new Map<string, boolean>();
	private frame!: HTMLIFrameElement;
	private stage!: HTMLElement;
	private ready = false;
	private styleBtns = new Map<string, HTMLElement>();
	private orientBtns: { el: HTMLElement; landscape: boolean }[] = [];

	constructor(
		app: App,
		private styles: PrintStyle[],
		prefer?: string
	) {
		super(app);
		this.styleId = this.styles.some((s) => s.id === prefer) ? (prefer as string) : this.styles[0].id;
		this.restore();
	}

	/* Kept where window positions are kept: which paper a printer likes is a
	   fact about this desk, not about the vault. */
	private store() {
		return this.app as unknown as { loadLocalStorage: (k: string) => unknown; saveLocalStorage: (k: string, v: unknown) => void };
	}

	private restore() {
		const raw = this.store().loadLocalStorage("powerdesk:print-opts") as { scale?: string; orient?: Record<string, boolean> } | null;
		if (!raw) return;
		if (PRINT_SCALES.some((s) => s.id === raw.scale)) this.scale = raw.scale as PrintScaleId;
		for (const [id, v] of Object.entries(raw.orient ?? {})) if (typeof v === "boolean") this.orient.set(id, v);
	}

	private remember() {
		this.store().saveLocalStorage("powerdesk:print-opts", { scale: this.scale, orient: Object.fromEntries(this.orient) });
	}

	private style(): PrintStyle {
		return this.styles.find((s) => s.id === this.styleId) ?? this.styles[0];
	}

	private options(): PrintOptions {
		const st = this.style();
		return { fontPt: scaledPt(st.basePt, this.scale), landscape: this.orient.get(st.id) ?? st.landscape };
	}

	onOpen() {
		this.titleEl.setText("Print");
		makeMovable(this.app, this, "powerdesk:print-window", { w: 920, h: 660 });
		const c = this.contentEl;
		c.addClass("pcal-print-wrap");
		const split = c.createDiv("pcal-print");

		const side = split.createDiv("pcal-print-side");

		side.createDiv({ cls: "pcal-print-label", text: "Style" });
		const list = side.createDiv("pcal-print-styles");
		for (const st of this.styles) {
			const b = list.createEl("button", { attr: { "aria-label": st.hint ?? st.label } });
			setIcon(b.createSpan("pcal-print-styleicon"), st.icon);
			b.createSpan({ text: st.label });
			if (st.hint) b.createSpan({ cls: "pcal-print-stylehint", text: st.hint });
			b.addEventListener("click", () => {
				this.styleId = st.id;
				this.sync();
			});
			this.styleBtns.set(st.id, b);
		}

		side.createDiv({ cls: "pcal-print-label", text: "Size" });
		const size = side.createEl("select", { cls: "dropdown" });
		for (const s of PRINT_SCALES) size.createEl("option", { value: s.id, text: s.label });
		size.value = this.scale;
		size.addEventListener("change", () => {
			this.scale = size.value as PrintScaleId;
			this.sync();
		});

		side.createDiv({ cls: "pcal-print-label", text: "Paper" });
		const orient = side.createDiv("pcal-print-orient");
		for (const [label, land] of [
			["Portrait", false],
			["Landscape", true],
		] as [string, boolean][]) {
			const b = orient.createEl("button", { attr: { "aria-label": label } });
			// the shape itself rather than an icon of one: two rectangles say
			// which way up the paper goes better than any glyph, and neither
			// depends on which Lucide set this Obsidian happens to ship
			b.createSpan({ cls: `pcal-print-paper${land ? " is-wide" : ""}` });
			b.createSpan({ text: label });
			b.addEventListener("click", () => {
				this.orient.set(this.style().id, land);
				this.sync();
			});
			this.orientBtns.push({ el: b, landscape: land });
		}

		side.createDiv({
			cls: "pcal-print-note",
			text: "The preview is the document. Print sends exactly this to the printer, where the paper size and the printer itself are chosen.",
		});

		this.stage = split.createDiv("pcal-print-stage");
		this.frame = this.stage.createEl("iframe", { cls: "pcal-print-preview", attr: { title: "Print preview" } });
		// registered before anything can ask to print, so the flag is already
		// true by the time a waiting print is woken by the same event
		this.frame.addEventListener("load", () => {
			this.ready = true;
		});

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createEl("button", { text: "Print", cls: "mod-cta pcal-send-btn" }).addEventListener("click", () => this.print());
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());

		this.sync();
	}

	/** Redraw the preview and put every control in step with it. */
	private sync() {
		const o = this.options();
		this.ready = false;
		this.frame.srcdoc = this.style().build(o);
		this.stage.toggleClass("is-landscape", !!o.landscape);
		for (const [id, b] of this.styleBtns) b.toggleClass("is-active", id === this.styleId);
		for (const b of this.orientBtns) b.el.toggleClass("is-active", b.landscape === !!o.landscape);
		this.remember();
	}

	/** Print the frame that is on screen.
	 *
	 *  The window stays open afterwards: the system dialog is where printing
	 *  is cancelled, and there is no way to tell a cancel from a print, so
	 *  closing here would throw away a window the user may still be using. */
	private print() {
		const w = this.frame.contentWindow;
		if (!w) return;
		// a document set through srcdoc arrives a tick later, and its images a
		// tick after that; printing before then prints a blank sheet, so a
		// click that beats the render waits for it rather than wasting paper
		if (!this.ready) {
			this.frame.addEventListener("load", () => this.print(), { once: true });
			return;
		}
		w.focus();
		w.print();
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Put a node in at the caret, or at the end when the caret is somewhere else
 *  entirely (the file picker that leads here can take focus away).
 *
 *  This is what execCommand("insertHTML") did, minus the string: that call took
 *  markup as text, so an image whose filename held a quote wrote a broken tag
 *  into the message. The command is also deprecated, and a range and a node say
 *  the same thing in the API that replaced it. */
function insertAtCaret(editor: HTMLElement, node: Node) {
	editor.focus();
	const sel = window.getSelection();
	const range = sel?.rangeCount ? sel.getRangeAt(0) : null;
	if (!sel || !range || !editor.contains(range.commonAncestorContainer)) {
		editor.appendChild(node);
		return;
	}
	range.deleteContents();
	range.insertNode(node);
	// leave the caret after what was just inserted, so typing carries on there
	range.setStartAfter(node);
	range.collapse(true);
	sel.removeAllRanges();
	sel.addRange(range);
}

/** The formatting bar shared by the compose window and the signature editor.
 *
 *  The buttons swallow mousedown so the editor keeps its selection: without
 *  that, clicking Bold moves focus to the button and there is nothing
 *  selected left to embolden. The editor is passed as a getter because the
 *  compose window builds its bar before its editor exists.
 *
 *  Bold and its neighbours are still execCommand: it is deprecated, and every
 *  browser still implements it, because nothing has replaced it. Doing this
 *  properly means a rich-text engine of our own (toggling a run that is half
 *  bold already, splitting a list item at the caret, undo that matches the
 *  editor's), which is a large thing to get wrong in a box people write mail
 *  in. The two calls that could leave, image insertion, have. */
function richToolbar(app: App, bar: HTMLElement, editor: () => HTMLElement, extra?: { label: string; icon: string; run: () => void }[]) {
	const tb = (icon: string, label: string, run: () => void) => {
		const b = bar.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
		setIcon(b, icon);
		b.addEventListener("mousedown", (e) => e.preventDefault());
		b.addEventListener("click", () => {
			run();
			editor().focus();
		});
	};
	const cmd = (name: string, arg?: string) => () => document.execCommand(name, false, arg);
	tb("bold", "Bold", cmd("bold"));
	tb("italic", "Italic", cmd("italic"));
	tb("underline", "Underline", cmd("underline"));
	tb("strikethrough", "Strikethrough", cmd("strikeThrough"));
	tb("list", "Bulleted list", cmd("insertUnorderedList"));
	tb("list-ordered", "Numbered list", cmd("insertOrderedList"));
	tb("link", "Link", () => {
		const sel = window.getSelection();
		const range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
		new PromptModal(app, "Link", [{ label: "Address", value: "", placeholder: "https://..." }], ([url]) => {
			if (!url.trim()) return;
			editor().focus();
			if (range && sel) {
				sel.removeAllRanges();
				sel.addRange(range);
			}
			document.execCommand("createLink", false, url.trim());
		}).open();
	});
	tb("eraser", "Clear formatting", cmd("removeFormat"));
	for (const x of extra ?? []) tb(x.icon, x.label, x.run);
}

/** Outlook's own text colors, enough to mark something up without turning
 *  the bar into a paint program. */
// prettier-ignore
const TEXT_COLORS: [string, string][] = [
	["Black", "#000000"], ["Dark grey", "#404040"], ["Grey", "#808080"],
	["Red", "#c00000"], ["Orange", "#e36c0a"], ["Yellow", "#bf8f00"],
	["Green", "#217346"], ["Teal", "#0f6c6c"], ["Blue", "#0f6cbd"],
	["Navy", "#1f3864"], ["Purple", "#7030a0"], ["Pink", "#c00060"],
];
// prettier-ignore
const HIGHLIGHTS: [string, string][] = [
	["Yellow", "#ffff00"], ["Green", "#a9f5a9"], ["Cyan", "#a9e9ff"],
	["Pink", "#ffb6e1"], ["Grey", "#d9d9d9"], ["None", "transparent"],
];

/** The fuller bar the compose window gets: sizes, colors, alignment, and the
 *  block-level bits Outlook puts on its Message tab. Kept apart from the
 *  basic bar so the signature editor is not asked to be a word processor. */
function richToolbarFull(app: App, bar: HTMLElement, editor: () => HTMLElement) {
	const btn = (icon: string, label: string, run: (e: MouseEvent) => void) => {
		const b = bar.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
		setIcon(b, icon);
		b.addEventListener("mousedown", (e) => e.preventDefault());
		b.addEventListener("click", (e) => {
			run(e);
			editor().focus();
		});
		return b;
	};
	const cmd = (name: string, arg?: string) => document.execCommand(name, false, arg);

	// size, as the four steps anyone actually uses
	const sizes: [string, string][] = [
		["Small", "2"],
		["Normal", "3"],
		["Large", "5"],
		["Huge", "6"],
	];
	btn("type", "Text size", (e) => {
		const menu = new Menu();
		for (const [name, v] of sizes) menu.addItem((i) => i.setTitle(name).onClick(() => cmd("fontSize", v)));
		menu.showAtMouseEvent(e);
	});
	const swatchMenu = (e: MouseEvent, colors: [string, string][], apply: (hex: string) => void) => {
		const menu = new Menu();
		for (const [name, hex] of colors)
			menu.addItem((i) => {
				i.setTitle(name).onClick(() => apply(hex));
				const dom = (i as unknown as { dom?: HTMLElement }).dom;
				if (dom) {
					const dot = dom.createSpan("pcal-cat-swatch");
					dot.style.backgroundColor = hex === "transparent" ? "var(--background-modifier-border)" : hex;
					dom.prepend(dot);
				}
			});
		menu.showAtMouseEvent(e);
	};
	btn("baseline", "Text color", (e) => swatchMenu(e, TEXT_COLORS, (hex) => cmd("foreColor", hex)));
	btn("highlighter", "Highlight", (e) => swatchMenu(e, HIGHLIGHTS, (hex) => cmd("hiliteColor", hex === "transparent" ? "inherit" : hex)));
	btn("align-left", "Align left", () => cmd("justifyLeft"));
	btn("align-center", "Center", () => cmd("justifyCenter"));
	btn("align-right", "Align right", () => cmd("justifyRight"));
	btn("indent", "Indent", () => cmd("indent"));
	btn("outdent", "Outdent", () => cmd("outdent"));
	btn("quote", "Quote", () => cmd("formatBlock", "blockquote"));
	btn("minus", "Divider", () => cmd("insertHorizontalRule"));
}

/** Let a modal be dragged by its title and resized from its corner, and have
 *  it open where it was left.
 *
 *  The geometry is kept in Obsidian's per-device local storage rather than in
 *  the vault's settings: where a window sits is a fact about this screen, and
 *  syncing it would drag a position from a 34-inch monitor onto a laptop and
 *  put the window half off the edge. Restoring clamps to the current viewport
 *  for the same reason, since screens change.
 */
function makeMovable(app: App, modal: Modal, key: string, fallback: { w: number; h: number }) {
	const stop = makeDraggable(app, modal.modalEl, modal.titleEl, key, fallback);
	const close = modal.onClose.bind(modal);
	modal.onClose = () => {
		stop();
		close();
	};
}

/** The moving and resizing itself, for anything on screen rather than only a
 *  Modal: the event card is a floating panel, not a dialog, and wants the
 *  same behavior. Returns the teardown, which also saves. */
function makeDraggable(
	app: App,
	el: HTMLElement,
	grip: HTMLElement,
	key: string,
	fallback: { w: number; h: number },
	/** Where to put it when nothing has been remembered yet. Without one it
	 *  centers, which is right for a dialog and wrong for a peek. */
	anchor?: () => { x: number; y: number }
): () => void {
	const store = app as unknown as { loadLocalStorage: (k: string) => unknown; saveLocalStorage: (k: string, v: unknown) => void };
	el.addClass("pcal-movable");

	const saved = store.loadLocalStorage(key) as { x?: number; y?: number; w?: number; h?: number } | null;
	const w = Math.min(Math.max(saved?.w ?? fallback.w, 380), window.innerWidth - 20);
	const h = Math.min(Math.max(saved?.h ?? fallback.h, 300), window.innerHeight - 20);
	el.style.width = `${w}px`;
	el.style.height = `${h}px`;
	// a window remembered off the edge of a smaller screen is a lost window
	const home = anchor?.() ?? { x: (window.innerWidth - w) / 2, y: (window.innerHeight - h) / 2 };
	const x = Math.min(Math.max(saved?.x ?? home.x, 0), Math.max(0, window.innerWidth - w));
	const y = Math.min(Math.max(saved?.y ?? home.y, 0), Math.max(0, window.innerHeight - h));
	el.style.left = `${x}px`;
	el.style.top = `${y}px`;

	// Position is kept from the first time onwards, like a dialog: an anchor
	// only decides where something lands when nothing has been remembered
	// yet. Anything that wants to go back to anchoring clears the stored x
	// and y, which is what the event card's "Follow the event again" does.
	const remember = () => store.saveLocalStorage(key, { x: el.offsetLeft, y: el.offsetTop, w: el.offsetWidth, h: el.offsetHeight });

	// dragging by the grip
	const bar = grip;
	bar.addClass("pcal-movable-grip");
	bar.addEventListener("mousedown", (e) => {
		if (e.button !== 0) return;
		e.preventDefault();
		const startX = e.clientX - el.offsetLeft;
		const startY = e.clientY - el.offsetTop;
		const move = (ev: MouseEvent) => {
			el.style.left = `${Math.min(Math.max(ev.clientX - startX, 0), window.innerWidth - el.offsetWidth)}px`;
			el.style.top = `${Math.min(Math.max(ev.clientY - startY, 0), window.innerHeight - el.offsetHeight)}px`;
		};
		const up = () => {
			document.removeEventListener("mousemove", move);
			document.removeEventListener("mouseup", up);
			remember();
		};
		document.addEventListener("mousemove", move);
		document.addEventListener("mouseup", up);
	});

	// the resize corner is the browser's own, so only the saving is ours
	let last = `${el.offsetWidth}x${el.offsetHeight}`;
	const watch = window.setInterval(() => {
		const now = `${el.offsetWidth}x${el.offsetHeight}`;
		if (now === last) return;
		last = now;
		remember();
	}, 500);
	return () => {
		window.clearInterval(watch);
		remember();
	};
}

/** Recipient autocomplete on a comma-separated address box.
 *
 *  Attaches to an input and owns a floating list under it. Only the fragment
 *  after the last comma is ever matched or replaced, so addresses already
 *  entered are never disturbed. The list is built per keystroke off an index
 *  that is already in memory, which is cheap enough not to need debouncing.
 *
 *  Keyboard first, because a recipient box is somewhere hands never leave the
 *  keyboard: arrows move, Enter and Tab take the highlighted one, Escape
 *  dismisses without taking anything. Enter with nothing highlighted submits
 *  the way it always did. */
class AddressSuggest {
	private listEl: HTMLElement | null = null;
	private hits: ContactHit[] = [];
	private active = -1;

	constructor(
		private input: HTMLInputElement,
		private plugin: PowerDeskPlugin
	) {
		void this.plugin.ensureSentContacts();
		input.setAttribute("autocomplete", "off");
		input.addEventListener("input", () => this.refresh());
		input.addEventListener("focus", () => this.refresh());
		input.addEventListener("blur", () => window.setTimeout(() => this.hide(), 150));
		input.addEventListener("keydown", (e) => this.onKey(e));
	}

	/** True when the list took the key, so the caller leaves it alone. */
	private onKey(e: KeyboardEvent): void {
		if (!this.listEl || !this.hits.length) return;
		if (e.key === "ArrowDown" || e.key === "ArrowUp") {
			e.preventDefault();
			const n = this.hits.length;
			this.active = e.key === "ArrowDown" ? (this.active + 1) % n : (this.active - 1 + n) % n;
			this.paint();
			return;
		}
		if (e.key === "Escape") {
			e.preventDefault();
			e.stopPropagation();
			this.hide();
			return;
		}
		if ((e.key === "Enter" || e.key === "Tab") && this.active >= 0) {
			e.preventDefault();
			e.stopPropagation();
			this.take(this.hits[this.active]);
		}
	}

	private refresh() {
		const caret = this.input.selectionStart ?? this.input.value.length;
		const frag = currentAddressFragment(this.input.value, caret);
		// an empty box offers the top of the list; an empty fragment after a
		// comma does not, since that is someone mid-thought, not asking
		if (!frag.text && this.input.value.trim()) return this.hide();
		this.hits = matchContacts(this.plugin.contactIndex(), frag.text);
		if (!this.hits.length) return this.hide();
		this.active = 0;
		this.paint();
	}

	private paint() {
		if (!this.listEl) {
			this.listEl = document.body.createDiv("pcal-addr-suggest");
			this.listEl.addEventListener("mousedown", (e) => e.preventDefault());
		}
		const el = this.listEl;
		el.empty();
		const r = this.input.getBoundingClientRect();
		el.style.left = `${r.left}px`;
		el.style.top = `${r.bottom + 2}px`;
		el.style.width = `${r.width}px`;
		this.hits.forEach((c, i) => {
			const row = el.createDiv("pcal-addr-row");
			row.toggleClass("is-active", i === this.active);
			row.createSpan({ cls: "pcal-addr-name", text: c.name || c.email });
			if (c.name) row.createSpan({ cls: "pcal-addr-email", text: c.email });
			row.addEventListener("click", () => this.take(c));
		});
	}

	private take(c: ContactHit) {
		const caret = this.input.selectionStart ?? this.input.value.length;
		const next = applyAddressChoice(this.input.value, caret, c.email);
		this.input.value = next.value;
		this.input.setSelectionRange(next.caret, next.caret);
		this.hide();
		this.input.focus();
	}

	hide() {
		this.listEl?.remove();
		this.listEl = null;
		this.active = -1;
		this.hits = [];
	}
}

/** Pick a time: the presets as buttons, and a real picker under them for
 *  anything they do not cover. Shared by snooze and schedule send so the two
 *  offer the same choices in the same order. */
class WhenModal extends Modal {
	constructor(
		app: App,
		private heading: string,
		private note: string,
		private use24h: boolean,
		private onPick: (ms: number) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText(this.heading);
		const c = this.contentEl;
		c.addClass("pcal-when");
		const now = Date.now();
		for (const p of whenPresets(now)) {
			const row = c.createEl("button", { cls: "pcal-when-row" });
			row.createSpan({ cls: "pcal-when-label", text: p.label });
			row.createSpan({ cls: "pcal-when-time", text: fmtWhen(p.ms, this.use24h) });
			row.addEventListener("click", () => {
				this.close();
				this.onPick(p.ms);
			});
		}
		const custom = c.createDiv("pcal-when-custom");
		custom.createDiv({ cls: "pcal-when-head", text: "Or pick a time" });
		const input = custom.createEl("input", { attr: { type: "datetime-local" } });
		// default the picker an hour out, rounded, so it opens somewhere sane
		const soon = new Date(now + 3600000);
		soon.setMinutes(soon.getMinutes() < 30 ? 0 : 30, 0, 0);
		const pad = (n: number) => String(n).padStart(2, "0");
		input.value = `${soon.getFullYear()}-${pad(soon.getMonth() + 1)}-${pad(soon.getDate())}T${pad(soon.getHours())}:${pad(soon.getMinutes())}`;
		const go = custom.createEl("button", { text: "Set", cls: "mod-cta" });
		const submit = () => {
			const ms = new Date(input.value).getTime();
			if (!Number.isFinite(ms)) {
				new Notice("Power Desk: that is not a time.");
				return;
			}
			if (ms <= Date.now()) {
				new Notice("Power Desk: pick a time in the future.");
				return;
			}
			this.close();
			this.onPick(ms);
		};
		go.addEventListener("click", submit);
		input.addEventListener("keydown", (e) => {
			if (e.key === "Enter") submit();
		});
		if (this.note) c.createDiv({ cls: "pcal-when-note", text: this.note });
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Any file in the vault, for attaching. Every file rather than just notes:
 *  the thing worth sending is usually a PDF or an image sitting in an
 *  attachments folder, not the note beside it. */
class VaultFilePickModal extends FuzzySuggestModal<TFile> {
	constructor(
		app: App,
		private onPick: (f: TFile) => void
	) {
		super(app);
		this.setPlaceholder("Attach a file from the vault...");
	}

	getItems(): TFile[] {
		return this.app.vault.getFiles().sort((a, b) => b.stat.mtime - a.stat.mtime);
	}

	getItemText(f: TFile): string {
		return f.path;
	}

	renderSuggestion(m: { item: TFile }, el: HTMLElement): void {
		el.addClass("pcal-palette-row");
		el.createSpan({ cls: "pcal-palette-label", text: m.item.path });
		el.createSpan({ cls: "pcal-palette-hint", text: fmtAttachmentSize(m.item.stat.size) });
	}

	onChooseItem(f: TFile): void {
		this.onPick(f);
	}
}

/** Signatures, laid out the way Outlook lays them out: the list on the left,
 *  the editor under it, and which one each account uses for new mail and for
 *  replies at the bottom.
 *
 *  Rich rather than a text box, because a real signature has a bold name, a
 *  colored title, and a logo. Images are embedded as data URLs so a
 *  signature is one self-contained thing that survives being copied between
 *  machines; they are turned into proper inline attachments at send time. */
class SignaturesModal extends Modal {
	private editing: string | null = null;
	private editorEl!: HTMLElement;
	private accountId: string;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
		this.accountId = plugin.mailAccounts()[0]?.id ?? "";
		this.editing = plugin.settings.mailSignatures[0]?.id ?? null;
	}

	onOpen() {
		this.titleEl.setText("Signatures");
		this.modalEl.addClass("pcal-sig-window");
		makeMovable(this.app, this, "powerdesk:signatures-window", { w: 760, h: 640 });
		this.draw();
	}

	private save() {
		this.plugin.queueSave();
	}

	/** Keep whatever is in the editor before the view changes under it. */
	private commit() {
		const sig = this.plugin.settings.mailSignatures.find((s) => s.id === this.editing);
		if (sig && this.editorEl?.isConnected) sig.html = this.editorEl.innerHTML;
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-sig");
		const sigs = this.plugin.settings.mailSignatures;

		// the list of signatures, and what can be done to one
		const top = c.createDiv("pcal-sig-top");
		const list = top.createDiv("pcal-sig-list");
		if (!sigs.length) list.createDiv({ cls: "pcal-when-note", text: "No signatures yet." });
		for (const s of sigs) {
			const row = list.createDiv("pcal-sig-item");
			row.toggleClass("is-current", s.id === this.editing);
			row.createSpan({ text: s.name });
			row.addEventListener("click", () => {
				this.commit();
				this.editing = s.id;
				this.draw();
			});
		}
		const side = top.createDiv("pcal-sig-side");
		side.createEl("button", { text: "New" }).addEventListener("click", () => {
			this.commit();
			new PromptModal(this.app, "New signature", [{ label: "Name", value: "", placeholder: "Work" }], ([name]) => {
				if (!name.trim()) return;
				const sig = { id: freshId(), name: name.trim(), html: "" };
				this.plugin.settings.mailSignatures = [...this.plugin.settings.mailSignatures, sig];
				this.editing = sig.id;
				this.save();
				this.draw();
			}).open();
		});
		const cur = sigs.find((s) => s.id === this.editing);
		const rename = side.createEl("button", { text: "Rename" });
		rename.disabled = !cur;
		rename.addEventListener("click", () => {
			if (!cur) return;
			this.commit();
			new PromptModal(this.app, "Rename signature", [{ label: "Name", value: cur.name }], ([name]) => {
				if (!name.trim()) return;
				cur.name = name.trim();
				this.save();
				this.draw();
			}).open();
		});
		const del = side.createEl("button", { text: "Delete" });
		del.disabled = !cur;
		del.addEventListener("click", () => {
			if (!cur) return;
			new ConfirmModal(this.app, `Delete the ${cur.name} signature?`, "Any account set to use it falls back to no signature.", "Delete", () => {
				this.plugin.settings.mailSignatures = this.plugin.settings.mailSignatures.filter((s) => s.id !== cur.id);
				// an account pointing at it must not silently take another
				for (const u of this.plugin.settings.mailSignatureUse) {
					if (u.newId === cur.id) u.newId = "";
					if (u.replyId === cur.id) u.replyId = "";
				}
				this.editing = this.plugin.settings.mailSignatures[0]?.id ?? null;
				this.save();
				this.draw();
			}).open();
		});

		// the editor
		if (cur) {
			const bar = c.createDiv("pcal-compose-bar");
			this.editorEl = c.createDiv({ cls: "pcal-compose-editor pcal-sig-editor", attr: { contenteditable: "true" } });
			richToolbar(this.app, bar, () => this.editorEl, [
				{
					label: "Insert an image",
					icon: "image",
					run: () =>
						new VaultFilePickModal(this.app, (f) => {
							void (async () => {
								try {
									const bytes = await this.app.vault.readBinary(f);
									const mime = mimeForExtension(f.extension);
									if (!mime.startsWith("image/")) {
										new Notice("Power Desk: pick an image file.");
										return;
									}
									const url = `data:${mime};base64,${arrayBufferToBase64(bytes)}`;
									insertAtCaret(this.editorEl, createEl("img", { attr: { src: url, alt: f.basename } }));
								} catch (err) {
									new Notice("Power Desk: could not read that image. " + (err instanceof Error ? err.message : String(err)));
								}
							})();
						}).open(),
				},
			]);
			this.editorEl.appendChild(sanitizeHTMLToDom(cur.html || "<p><br></p>"));
			this.editorEl.addEventListener("blur", () => this.commit());
		}

		// which signature this account uses, and when
		const accounts = this.plugin.mailAccounts();
		if (accounts.length) {
			c.createDiv({ cls: "pcal-shortcuts-head", text: "Use" });
			if (accounts.length > 1) {
				new Setting(c).setName("Account").addDropdown((d) => {
					for (const a of accounts) d.addOption(a.id, this.plugin.nameOf(a));
					d.setValue(this.accountId).onChange((v) => {
						this.commit();
						this.accountId = v;
						this.draw();
					});
				});
			}
			const use = this.plugin.signatureUseFor(this.accountId);
			const pick = (label: string, desc: string, get: () => string, set: (v: string) => void) =>
				new Setting(c)
					.setName(label)
					.setDesc(desc)
					.addDropdown((d) => {
						d.addOption("", "None");
						for (const s of this.plugin.settings.mailSignatures) d.addOption(s.id, s.name);
						d.setValue(get()).onChange((v) => {
							set(v);
							this.save();
						});
					});
			pick("New messages", "Put on a message you start.", () => use.newId, (v) => (use.newId = v));
			pick("Replies and forwards", "Often shorter than the one on a first message.", () => use.replyId, (v) => (use.replyId = v));
		}

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => this.close());
	}

	onClose() {
		this.commit();
		this.save();
		this.contentEl.empty();
	}
}

/** The mailbox's category list: make one, recolor one, delete one.
 *
 *  No rename, because Graph has none. Its own documentation is blunt about
 *  it: the display name cannot be modified once a category exists. Outlook
 *  appears to rename because it rewrites every item carrying the old name at
 *  the same time, which is a mailbox-wide rewrite this plugin is not going
 *  to do behind a button labelled "rename". The limit is stated in the UI
 *  rather than hidden behind a control that would fail. */
class CategoriesModal extends Modal {
	private accountId: string;
	private loading = true;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onChange: () => void
	) {
		super(app);
		this.accountId = plugin.mailAccounts()[0]?.id ?? "";
	}

	onOpen() {
		this.titleEl.setText("Categories");
		makeMovable(this.app, this, "powerdesk:categories-window", { w: 620, h: 540 });
		this.draw();
		void this.load();
	}

	private async load() {
		this.loading = true;
		this.draw();
		await this.plugin.loadCategories(this.accountId);
		this.loading = false;
		this.draw();
	}

	private colorMenu(e: MouseEvent, current: string, pick: (preset: string) => void) {
		const menu = new Menu();
		for (let i = 0; i < 25; i++) {
			const preset = `preset${i}`;
			menu.addItem((it) => {
				it.setTitle(CATEGORY_COLOR_NAMES[i] ?? preset)
					.setChecked(current === preset)
					.onClick(() => pick(preset));
				// a swatch beside the name, since the names alone are a quiz
				const dom = (it as unknown as { dom?: HTMLElement }).dom;
				if (dom) {
					const dot = dom.createSpan("pcal-cat-swatch");
					dot.style.backgroundColor = categoryColor(preset);
					dom.prepend(dot);
				}
			});
		}
		menu.showAtMouseEvent(e);
	}

	/** The nearest thing to a rename: ask for the new name, count what
	 *  carries the old one, say exactly what will be rewritten and what
	 *  cannot be reached, and only then do it. */
	private askReplace(cat: OutlookCategory) {
		new PromptModal(this.app, `Replace "${cat.displayName}" with`, [{ label: "New name", value: cat.displayName }], ([name]) => {
			const to = name.trim();
			if (!to || to.toLowerCase() === cat.displayName.toLowerCase()) return;
			const counting = new Notice(`Power Desk: finding everything tagged ${cat.displayName}...`, 0);
			void this.plugin.countCategoryUse(this.accountId, cat.displayName).then((found) => {
				counting.hide();
				if (!found) return;
				const n = found.hits.length;
				const capped = found.complete ? "" : ` Only the first ${n} are covered, so run it again afterwards to catch the rest.`;
				new ConfirmModal(
					this.app,
					`Replace "${cat.displayName}" with "${to}" on ${n} message${n === 1 ? "" : "s"}?`,
					`Power Desk will make "${to}" in the same color, retag ${n === 1 ? "that message" : `those ${n} messages`}, and then retire "${cat.displayName}". This reaches mail only: calendar events, tasks, and contacts carrying the old name keep it, and you would have to change those in Outlook. If any message cannot be rewritten the old category is left in place, so nothing ends up orphaned.${capped}`,
					n ? `Replace on ${n}` : "Replace",
					() => {
						const bar = new Notice(`Power Desk: replacing 0 of ${n}...`, 0);
						void this.plugin
							.replaceCategoryEverywhere(this.accountId, cat, to, found.hits, (done, total) => bar.setMessage(`Power Desk: replacing ${done} of ${total}...`))
							.then((r) => {
								bar.hide();
								if (r.failed) new Notice(`Power Desk: retagged ${r.changed}, but ${r.failed} could not be rewritten, so "${cat.displayName}" was kept. Run it again to finish.`, 12000);
								else new Notice(r.retired ? `Power Desk: "${cat.displayName}" is now "${to}" on ${r.changed} message${r.changed === 1 ? "" : "s"}.` : `Power Desk: retagged ${r.changed}, but the old category could not be removed.`, 9000);
								this.draw();
								this.onChange();
							});
					}
				).open();
			});
		}).open();
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-cats");
		const accounts = this.plugin.mailAccounts();
		if (accounts.length > 1) {
			new Setting(c).setName("Account").addDropdown((d) => {
				for (const a of accounts) d.addOption(a.id, this.plugin.nameOf(a));
				d.setValue(this.accountId).onChange((v) => {
					this.accountId = v;
					void this.load();
				});
			});
		}
		const acc = this.plugin.accountById(this.accountId);
		if (acc && !acc.grantedScope.includes("MailboxSettings")) {
			c.createDiv({ cls: "pcal-mail-error", text: `Reconnect ${this.plugin.nameOf(acc)} in settings to let Power Desk manage its categories.` });
			return;
		}
		if (this.loading) {
			c.createDiv({ cls: "pcal-when-note", text: "Reading the mailbox..." });
			return;
		}

		const cats = this.plugin.categoriesFor(this.accountId);
		if (!cats.length) c.createDiv({ cls: "pcal-when-note", text: "No categories yet." });
		for (const cat of cats) {
			const row = c.createDiv("pcal-cat-row");
			const dot = row.createSpan("pcal-cat-swatch");
			dot.style.backgroundColor = categoryColor(cat.color);
			row.createSpan({ cls: "pcal-cat-name", text: cat.displayName });
			row.createEl("button", { text: "Color" }).addEventListener("click", (e) => this.colorMenu(e, cat.color, (preset) => void this.plugin.recolorCategory(this.accountId, cat.id, preset).then(() => {
				this.draw();
				this.onChange();
			})));
			row.createEl("button", { text: "Replace everywhere" }).addEventListener("click", () => this.askReplace(cat));
			row.createEl("button", { text: "Delete" }).addEventListener("click", () => {
				new ConfirmModal(
					this.app,
					`Delete the ${cat.displayName} category?`,
					"It leaves the mailbox's list, so nothing new can be tagged with it. Messages already carrying it keep the label, but it loses its color and stops appearing in the Categorize menu.",
					"Delete category",
					() => {
						void this.plugin.removeCategory(this.accountId, cat.id).then(() => {
							this.draw();
							this.onChange();
						});
					}
				).open();
			});
		}

		c.createDiv({
			cls: "pcal-when-note",
			text: "A category's name cannot be changed once it is made: Graph allows only the color to be edited. Replace everywhere is the honest substitute, and it says so: it makes the new name in the same color, retags every message carrying the old one, and retires the old only if every rewrite worked. It reaches mail alone, so calendar events, tasks, and contacts keep the old name and need Outlook.",
		});

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "New category", cls: "mod-cta" }).addEventListener("click", () => {
			new PromptModal(this.app, "New category", [{ label: "Name", value: "", placeholder: "Waiting on" }], ([name]) => {
				if (!name.trim()) return;
				// a fresh category gets the first color nothing else is using,
				// so a new one is distinguishable without a second decision
				void this.plugin.newCategory(this.accountId, name, this.plugin.freeCategoryColor(this.accountId)).then(() => {
					this.draw();
					this.onChange();
				});
			}).open();
		});
		btns.createEl("button", { text: "Done" }).addEventListener("click", () => this.close());
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Out of office, per account.
 *
 *  One account at a time on purpose: the mailboxes are separate, the replies
 *  say different things, and a single switch that turned on both a work and a
 *  personal auto-reply would be a mistake waiting to be made. */
class OutOfOfficeModal extends Modal {
	private accountId: string;
	private draft: AutoReply = { status: "disabled", externalAudience: "all", internalReplyMessage: "", externalReplyMessage: "" };
	private startMs = Date.now();
	private endMs = Date.now() + 86400000;
	private loading = true;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
		this.accountId = plugin.mailAccounts()[0]?.id ?? "";
	}

	onOpen() {
		this.titleEl.setText("Automatic replies");
		// taller than it looks: two message boxes appear once it is switched on
		makeMovable(this.app, this, "powerdesk:ooo-window", { w: 660, h: 620 });
		void this.load();
		this.draw();
	}

	private async load() {
		this.loading = true;
		this.draw();
		const s = await this.plugin.loadAutoReply(this.accountId);
		if (s) {
			this.draft = s;
			this.startMs = fromGraphDateTime(s.scheduledStartDateTime) ?? Date.now();
			this.endMs = fromGraphDateTime(s.scheduledEndDateTime) ?? Date.now() + 86400000;
		} else {
			this.draft = { status: "disabled", externalAudience: "all", internalReplyMessage: "", externalReplyMessage: "" };
		}
		this.loading = false;
		this.draw();
	}

	private local(ms: number): string {
		const d = new Date(ms);
		const p = (n: number) => String(n).padStart(2, "0");
		return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-ooo");
		const accounts = this.plugin.mailAccounts();
		if (accounts.length > 1) {
			new Setting(c).setName("Account").addDropdown((d) => {
				for (const a of accounts) d.addOption(a.id, this.plugin.nameOf(a));
				d.setValue(this.accountId).onChange((v) => {
					this.accountId = v;
					void this.load();
				});
			});
		}
		const acc = this.plugin.accountById(this.accountId);
		if (acc && !acc.grantedScope.includes("MailboxSettings")) {
			c.createDiv({ cls: "pcal-mail-error", text: `Reconnect ${this.plugin.nameOf(acc)} in settings to let Power Desk read and set its automatic replies.` });
			return;
		}
		if (this.loading) {
			c.createDiv({ cls: "pcal-when-note", text: "Reading the mailbox..." });
			return;
		}

		new Setting(c).setName("Automatic replies").addDropdown((d) => {
			d.addOption("disabled", "Off");
			d.addOption("alwaysEnabled", "On until I turn it off");
			d.addOption("scheduled", "On for a period");
			d.setValue(this.draft.status).onChange((v) => {
				this.draft.status = v as AutoReply["status"];
				this.draw();
			});
		});

		if (this.draft.status === "scheduled") {
			const when = (label: string, get: () => number, set: (ms: number) => void) =>
				new Setting(c).setName(label).addText((t) => {
					t.inputEl.type = "datetime-local";
					t.setValue(this.local(get())).onChange((v) => {
						const ms = new Date(v).getTime();
						if (Number.isFinite(ms)) set(ms);
					});
				});
			when("Start", () => this.startMs, (ms) => (this.startMs = ms));
			when("End", () => this.endMs, (ms) => (this.endMs = ms));
		}

		if (this.draft.status !== "disabled") {
			const area = (label: string, desc: string, get: () => string, set: (v: string) => void) =>
				new Setting(c)
					.setName(label)
					.setDesc(desc)
					.addTextArea((t) => {
						t.setValue(get()).onChange(set);
						t.inputEl.rows = 5;
						t.inputEl.addClass("pcal-ooo-text");
					});
			area("Reply to colleagues", "Sent to people inside your organization.", () => this.draft.internalReplyMessage, (v) => (this.draft.internalReplyMessage = v));
			new Setting(c)
				.setName("Reply to outsiders")
				.setDesc("Who outside your organization gets a reply at all.")
				.addDropdown((d) => {
					d.addOption("none", "Nobody");
					d.addOption("contactsOnly", "My contacts only");
					d.addOption("all", "Everyone who writes");
					d.setValue(this.draft.externalAudience).onChange((v) => {
						this.draft.externalAudience = v as AutoReply["externalAudience"];
						this.draw();
					});
				});
			if (this.draft.externalAudience !== "none")
				area("Message to outsiders", "Sent to people outside your organization.", () => this.draft.externalReplyMessage, (v) => (this.draft.externalReplyMessage = v));
			c.createDiv({
				cls: "pcal-when-note",
				text: "Plain text is fine; paste HTML if you want formatting. The mailbox sends these, so they go out whether or not Obsidian is open, and they are the same replies Outlook shows.",
			});
		}

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			if (this.draft.status === "scheduled") {
				if (!(this.endMs > this.startMs)) {
					new Notice("Power Desk: the end has to come after the start.");
					return;
				}
				this.draft.scheduledStartDateTime = toGraphDateTime(this.startMs);
				this.draft.scheduledEndDateTime = toGraphDateTime(this.endMs);
			}
			void this.plugin.saveAutoReply(this.accountId, this.draft).then((ok) => {
				if (ok) this.close();
			});
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Choose the toolbar's buttons and their order.
 *
 *  Deliberately buttons rather than dragging: this is a list someone visits
 *  once, and up/down/remove works with a keyboard, on a phone, and for anyone
 *  who cannot drag. */
class ToolbarModal extends Modal {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private catalog: { id: string; label: string; icon: string }[],
		/** Which toolbar this is editing: read it, write it, and the order to
		 *  fall back to. Passed in rather than hardcoded, so mail and the
		 *  calendar share one editor instead of owning two. */
		private slot: { get: () => string[]; set: (v: string[]) => void; fallback: string[]; leads: string },
		private onChange: () => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Customize the toolbar");
		makeMovable(this.app, this, "powerdesk:toolbar-window", { w: 560, h: 620 });
		this.draw();
	}

	private save(next: string[]) {
		this.slot.set(next);
		this.plugin.queueSave();
		this.onChange();
		this.draw();
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-toolbar-edit");
		const shown = this.slot.get().filter((id) => this.catalog.some((a) => a.id === id));
		c.createDiv({
			cls: "pcal-when-note",
			text: `${this.slot.leads} always leads. Everything left off is still on a keyboard shortcut, the right-click menu, and the command palette, so a short toolbar gives nothing up.`,
		});

		c.createDiv({ cls: "pcal-shortcuts-head", text: "On the toolbar" });
		if (!shown.length) c.createDiv({ cls: "pcal-when-note", text: "Nothing yet." });
		shown.forEach((id, idx) => {
			const a = this.catalog.find((x) => x.id === id);
			if (!a) return;
			const row = c.createDiv("pcal-toolbar-row");
			setIcon(row.createSpan("pcal-toolbar-ic"), a.icon);
			row.createSpan({ cls: "pcal-toolbar-label", text: a.label });
			const move = (to: number) => {
				const next = [...shown];
				next.splice(to, 0, next.splice(idx, 1)[0]);
				this.save(next);
			};
			const up = row.createEl("button", { attr: { "aria-label": "Move up" } });
			setIcon(up, "chevron-up");
			up.disabled = idx === 0;
			up.addEventListener("click", () => move(idx - 1));
			const down = row.createEl("button", { attr: { "aria-label": "Move down" } });
			setIcon(down, "chevron-down");
			down.disabled = idx === shown.length - 1;
			down.addEventListener("click", () => move(idx + 1));
			const off = row.createEl("button", { attr: { "aria-label": `Remove ${a.label}` } });
			setIcon(off, "x");
			off.addEventListener("click", () => this.save(shown.filter((x) => x !== id)));
		});

		const rest = this.catalog.filter((a) => !shown.includes(a.id));
		if (rest.length) {
			c.createDiv({ cls: "pcal-shortcuts-head", text: "Available" });
			for (const a of rest) {
				const row = c.createDiv("pcal-toolbar-row");
				setIcon(row.createSpan("pcal-toolbar-ic"), a.icon);
				row.createSpan({ cls: "pcal-toolbar-label", text: a.label });
				const add = row.createEl("button", { attr: { "aria-label": `Add ${a.label}` } });
				setIcon(add, "plus");
				add.addEventListener("click", () => this.save([...shown, a.id]));
			}
		}

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Reset to default" }).addEventListener("click", () => this.save([...this.slot.fallback]));
		btns.createEl("button", { text: "Done", cls: "mod-cta" }).addEventListener("click", () => this.close());
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Outlook's Shortcuts, given something worth pointing at.
 *
 *  In Outlook a shortcut reaches a folder or a file. Here it can reach a mail
 *  folder, a saved search, a note in the vault, or a link, which is the one
 *  list neither application can offer on its own: the folder you file into
 *  and the note you write in it, side by side. */
class ShortcutsModal extends Modal {
	private listEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onFolder: (accountId: string, folderId: string, name: string) => void,
		private onSearch: (q: string) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Shortcuts");
		makeMovable(this.app, this, "powerdesk:shortcuts-window", { w: 620, h: 660 });
		const c = this.contentEl;
		c.addClass("pcal-shortcuts-modal");
		this.listEl = c.createDiv("pcal-notes-list");
		this.draw();

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createEl("button", { text: "Add", cls: "mod-cta pcal-send-btn" }).addEventListener("click", (e) => this.addMenu(e));
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private save() {
		this.plugin.queueSave();
		this.draw();
	}

	private add(kind: Shortcut["kind"], target: string, accountId?: string, label?: string) {
		const s: Shortcut = { id: freshId(), group: "", label: (label ?? "").trim() || defaultShortcutLabel(kind, target), kind, target, accountId };
		this.plugin.settings.shortcuts = [...this.plugin.settings.shortcuts, s];
		this.save();
	}

	private addMenu(e: MouseEvent) {
		const menu = new Menu();
		menu.addItem((i) =>
			i
				.setTitle("A mail folder...")
				.setIcon("folder")
				.onClick(() => {
					const targets = this.plugin.moveTargets([]);
					if (!targets.length) {
						new Notice("Power Desk: open the inbox once so the folders are known.");
						return;
					}
					new FolderPickModal(this.app, this.plugin, this.plugin.mailAccounts()[0]?.id ?? "", (folderId, name) => this.add("folder", folderId, this.plugin.mailAccounts()[0]?.id, name)).open();
				})
		);
		menu.addItem((i) =>
			i
				.setTitle("A saved search...")
				.setIcon("search")
				.onClick(() => {
					new PromptModal(this.app, "Shortcut to a search", [
						{ label: "Search", value: "", placeholder: "from:deanna is:unread" },
						{ label: "Name", value: "", placeholder: "optional" },
					], ([q, name]) => {
						if (q.trim()) this.add("search", q.trim(), undefined, name);
					}).open();
				})
		);
		menu.addItem((i) =>
			i
				.setTitle("A note...")
				.setIcon("file-text")
				.onClick(() => new VaultFilePickModal(this.app, (f) => this.add("note", f.path)).open())
		);
		menu.addItem((i) =>
			i
				.setTitle("A link...")
				.setIcon("link")
				.onClick(() => {
					new PromptModal(this.app, "Shortcut to a link", [
						{ label: "Address", value: "", placeholder: "https://..." },
						{ label: "Name", value: "", placeholder: "optional" },
					], ([url, name]) => {
						if (/^https?:\/\//i.test(url.trim())) this.add("url", url.trim(), undefined, name);
						else if (url.trim()) new Notice("Power Desk: a link shortcut needs an http or https address.");
					}).open();
				})
		);
		menu.showAtMouseEvent(e);
	}

	private follow(s: Shortcut) {
		if (s.kind === "folder" && s.accountId) this.onFolder(s.accountId, s.target, s.label);
		else if (s.kind === "search") this.onSearch(s.target);
		else if (s.kind === "note") {
			const f = this.app.vault.getAbstractFileByPath(s.target);
			if (f instanceof TFile) void this.plugin.showNote(f, false);
			else new Notice(`Power Desk: ${s.target} is not in the vault any more.`);
		} else if (s.kind === "url") window.open(s.target, "_blank");
		this.close();
	}

	private draw() {
		const host = this.listEl;
		host.empty();
		const all = this.plugin.settings.shortcuts;
		if (!all.length) {
			host.createDiv({ cls: "pcal-when-note", text: "No shortcuts yet. Add points at a mail folder, a saved search, a note, or a link." });
			return;
		}
		const icons: Record<Shortcut["kind"], string> = { folder: "folder", search: "search", note: "file-text", url: "link" };
		for (const g of groupShortcuts(all)) {
			if (g.name) host.createDiv({ cls: "pcal-shortcuts-head", text: g.name });
			for (const s of g.items) {
				const row = host.createDiv("pcal-folders-row");
				setIcon(row.createSpan("pcal-folder-ic"), icons[s.kind]);
				const mid = row.createDiv("pcal-folders-mid");
				mid.createDiv({ cls: "pcal-folders-name", text: s.label });
				mid.createDiv({ cls: "pcal-folders-counts", text: s.kind === "note" ? s.target : s.kind === "folder" ? "Mail folder" : s.target });
				const acts = row.createDiv("pcal-people-acts");
				const act = (icon: string, label: string, run: () => void) => {
					const b = acts.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
					setIcon(b, icon);
					b.addEventListener("click", (e) => {
						e.stopPropagation();
						run();
					});
				};
				const move = (n: number) => {
					const list = [...all];
					const i = list.findIndex((x) => x.id === s.id);
					const j = i + n;
					if (i < 0 || j < 0 || j >= list.length) return;
					list.splice(j, 0, list.splice(i, 1)[0]);
					this.plugin.settings.shortcuts = list;
					this.save();
				};
				act("chevron-up", "Move up", () => move(-1));
				act("chevron-down", "Move down", () => move(1));
				act("pencil", "Rename or group", () => {
					new PromptModal(this.app, "Shortcut", [
						{ label: "Name", value: s.label },
						{ label: "Group", value: s.group, placeholder: "optional" },
					], ([label, group]) => {
						s.label = label.trim() || s.label;
						s.group = group.trim();
						this.save();
					}).open();
				});
				act("x", "Remove", () => {
					this.plugin.settings.shortcuts = all.filter((x) => x.id !== s.id);
					this.save();
				});
				row.addEventListener("click", () => this.follow(s));
			}
		}
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Every folder in every account, with what is in it and what can be done
 *  to it.
 *
 *  The pane down the side is for getting to a folder; this is for looking
 *  after them: the whole tree at once including the hidden ones, counts,
 *  a filter for a mailbox with sixty folders, and the housekeeping that has
 *  nowhere else to live. */
class FoldersModal extends Modal {
	private filter = "";
	private listEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onOpenFolder: (accountId: string, folderId: string, name: string) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Folders");
		makeMovable(this.app, this, "powerdesk:folders-window", { w: 680, h: 700 });
		const c = this.contentEl;
		c.addClass("pcal-folders-modal");
		const search = c.createEl("input", { cls: "pcal-people-search", attr: { type: "search", placeholder: "Filter folders..." } });
		search.addEventListener("input", () => {
			this.filter = search.value.toLowerCase();
			this.draw();
		});
		this.listEl = c.createDiv("pcal-notes-list");
		this.draw();
		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private draw() {
		const host = this.listEl;
		host.empty();
		const accounts = this.plugin.mailAccounts();
		if (!accounts.length) {
			host.createDiv({ cls: "pcal-when-note", text: "No mail accounts connected." });
			return;
		}
		for (const a of accounts) {
			const tree = this.plugin.folderTreeFor(a);
			const hidden = this.plugin.settings.mailHiddenFolders.filter((h) => h.accountId === a.id);
			const rows = tree.filter((n) => !this.filter || n.folder.name.toLowerCase().includes(this.filter));
			host.createDiv({ cls: "pcal-shortcuts-head", text: this.plugin.nameOf(a) });
			if (!tree.length) {
				host.createDiv({ cls: "pcal-when-note", text: "Folders are still loading. Open the inbox once." });
				continue;
			}
			if (!rows.length) {
				host.createDiv({ cls: "pcal-when-note", text: "Nothing matches that." });
				continue;
			}
			const inboxId = this.plugin.inboxIdFor(a);
			for (const { folder, depth } of rows) {
				const row = host.createDiv("pcal-folders-row");
				// the filter flattens the tree, since indenting a filtered
				// list by a parent you cannot see is just a ragged left edge
				row.style.paddingLeft = `${8 + (this.filter ? 0 : depth * 14)}px`;
				const isHidden = hidden.some((h) => h.folderId === folder.id);
				row.toggleClass("is-hidden-folder", isHidden);
				setIcon(row.createSpan("pcal-folder-ic"), folder.id === inboxId ? "inbox" : "folder");
				const mid = row.createDiv("pcal-folders-mid");
				mid.createDiv({ cls: "pcal-folders-name", text: folder.name + (isHidden ? " (hidden)" : "") });
				mid.createDiv({ cls: "pcal-folders-counts", text: `${folder.total} item${folder.total === 1 ? "" : "s"}${folder.unread ? `, ${folder.unread} unread` : ""}` });
				const acts = row.createDiv("pcal-people-acts");
				const act = (icon: string, label: string, run: () => void, disabled = false) => {
					const b = acts.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
					setIcon(b, icon);
					b.disabled = disabled;
					b.addEventListener("click", (e) => {
						e.stopPropagation();
						run();
					});
				};
				act("mail-open", `Mark everything in ${folder.name} read`, () => this.markRead(a.id, folder), !folder.unread);
				act("folder-plus", "New subfolder", () => {
					new PromptModal(this.app, `New folder in ${folder.name}`, [{ label: "Name", value: "" }], ([name]) => {
						if (name.trim()) void this.plugin.newFolder(a.id, name, folder.id).then(() => this.draw());
					}).open();
				});
				const system = isSystemFolder(folder.name, folder.id, inboxId);
				act(
					"pencil",
					"Rename",
					() => {
						new PromptModal(this.app, `Rename ${folder.name}`, [{ label: "Name", value: folder.name }], ([name]) => {
							if (name.trim() && name.trim() !== folder.name) void this.plugin.renameFolder(a.id, folder.id, name).then(() => this.draw());
						}).open();
					},
					system
				);
				act(isHidden ? "eye" : "eye-off", isHidden ? "Show in the folder pane" : "Hide from the folder pane", () => {
					const s = this.plugin.settings;
					s.mailHiddenFolders = isHidden ? s.mailHiddenFolders.filter((h) => !(h.accountId === a.id && h.folderId === folder.id)) : [...s.mailHiddenFolders, { accountId: a.id, folderId: folder.id }];
					this.plugin.queueSave();
					this.plugin.notify();
					this.draw();
				});
				row.addEventListener("click", () => {
					this.onOpenFolder(a.id, folder.id, folder.name);
					this.close();
				});
			}
			const add = host.createDiv("pcal-modal-btns");
			add.createEl("button", { text: `New folder in ${this.plugin.nameOf(a)}` }).addEventListener("click", () => {
				new PromptModal(this.app, `New folder in ${this.plugin.nameOf(a)}`, [{ label: "Name", value: "" }], ([name]) => {
					if (name.trim()) void this.plugin.newFolder(a.id, name, null).then(() => this.draw());
				}).open();
			});
		}
	}

	private markRead(accountId: string, folder: MailFolder) {
		new ConfirmModal(
			this.app,
			`Mark everything in ${folder.name} as read?`,
			`${folder.unread} message${folder.unread === 1 ? "" : "s"} will be marked read. There is no bulk flag for this in the mailbox, so it is one write per message and a large folder takes a moment.`,
			"Mark all read",
			() => {
				const bar = new Notice(`Power Desk: marking 0 of ${folder.unread}...`, 0);
				void this.plugin
					.markFolderRead(accountId, folder.id, (done, total) => bar.setMessage(`Power Desk: marking ${done} of ${total}...`))
					.then((r) => {
						bar.hide();
						if (r.failed) new Notice(`Power Desk: marked ${r.done} read, ${r.failed} would not. Run it again to finish.`, 9000);
						else if (!r.complete) new Notice(`Power Desk: marked ${r.done} read. That folder holds more than one pass covers, so run it again.`, 9000);
						else if (r.done) new Notice(`Power Desk: marked ${r.done} message${r.done === 1 ? "" : "s"} read.`);
						this.draw();
					});
			}
		).open();
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** A day's activity, which is what Outlook's Journal recorded before it was
 *  retired: what you sat in, who wrote, what you sent.
 *
 *  Built from the mail and calendar rather than from the old Journal folder,
 *  which Microsoft stopped writing to years ago and which is empty in almost
 *  every mailbox now. What it does that Outlook's never could is put the day
 *  into your daily note, which is the reason to have it here. */
class JournalModal extends Modal {
	private key: string;
	private bodyEl!: HTMLElement;
	private loading = true;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
		this.key = keyOfDate(new Date());
	}

	onOpen() {
		this.titleEl.setText("Journal");
		makeMovable(this.app, this, "powerdesk:journal-window", { w: 660, h: 700 });
		const c = this.contentEl;
		c.addClass("pcal-journal");

		const head = c.createDiv("pcal-tasks-head");
		const back = head.createEl("button", { attr: { "aria-label": "Previous day" } });
		setIcon(back, "chevron-left");
		back.addEventListener("click", () => this.go(-1));
		const date = head.createEl("input", { attr: { type: "date" } });
		date.value = this.key;
		date.addEventListener("change", () => {
			if (/^\d{4}-\d{2}-\d{2}$/.test(date.value)) {
				this.key = date.value;
				void this.load();
			}
		});
		const fwd = head.createEl("button", { attr: { "aria-label": "Next day" } });
		setIcon(fwd, "chevron-right");
		fwd.addEventListener("click", () => this.go(1));
		head.createEl("button", { text: "Today" }).addEventListener("click", () => {
			this.key = keyOfDate(new Date());
			date.value = this.key;
			void this.load();
		});
		this.dateInput = date;

		this.bodyEl = c.createDiv("pcal-journal-body");
		void this.load();

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createEl("button", { text: "Add to the daily note", cls: "mod-cta pcal-send-btn" }).addEventListener("click", () => void this.toDailyNote());
		btns.createEl("button", { text: "Copy" }).addEventListener("click", () => {
			void navigator.clipboard.writeText(journalMarkdown(this.day())).then(() => new Notice("Power Desk: the day is on the clipboard."));
		});
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private dateInput!: HTMLInputElement;

	private go(n: number) {
		this.key = addDays(this.key, n);
		this.dateInput.value = this.key;
		void this.load();
	}

	private async load() {
		this.loading = true;
		this.draw();
		// the calendar and the sent folder both have to be there before the
		// day can be described
		this.plugin.ensureWindow(this.key, this.key, false);
		await this.plugin.ensureSentForJournal();
		this.loading = false;
		this.draw();
	}

	private day(): JournalDay {
		return buildJournal(
			this.key,
			this.plugin.eventsForWindow(this.key, this.key),
			this.plugin.allMail(),
			this.plugin.sentMail(),
			this.plugin.stickyNotes().map((n) => ({ title: n.title, changedMs: n.changedMs })),
			this.plugin.settings.use24h
		);
	}

	private draw() {
		const c = this.bodyEl;
		c.empty();
		const d = this.day();
		c.createDiv({ cls: "pcal-shortcuts-head", text: fmtDayHeading(this.key) });
		if (this.loading) c.createDiv({ cls: "pcal-when-note", text: "Reading the day..." });

		const section = (title: string, rows: { when?: string; main: string; sub?: string }[]) => {
			if (!rows.length) return;
			c.createDiv({ cls: "pcal-shortcuts-head", text: title });
			const box = c.createDiv("pcal-tasks-list");
			for (const r of rows) {
				const row = box.createDiv("pcal-tasks-row");
				if (r.when) row.createDiv({ cls: "pcal-eventsearch-when", text: r.when });
				const mid = row.createDiv("pcal-tasks-mid");
				mid.createDiv({ cls: "pcal-tasks-title", text: r.main });
				if (r.sub) mid.createDiv({ cls: "pcal-tasks-due", text: r.sub });
			}
		};
		section(
			"Meetings",
			d.meetings.map((m) => ({ when: m.when, main: m.title, sub: m.who }))
		);
		section(
			"Sent",
			d.sent.map((m) => ({ when: m.when, main: m.subject }))
		);
		section(
			"Received",
			d.received.map((m) => ({ when: m.when, main: m.subject, sub: m.who }))
		);
		section(
			"Notes",
			d.notes.map((n) => ({ main: n.title }))
		);
		if (!this.loading && !d.meetings.length && !d.sent.length && !d.received.length && !d.notes.length)
			c.createDiv({ cls: "pcal-when-note", text: "Nothing recorded for this day. Mail is only described as far back as it has been fetched." });
	}

	private async toDailyNote() {
		const ok = await this.plugin.appendToDailyNote(this.key, journalMarkdown(this.day()));
		if (ok) this.close();
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Outlook's notes, the module everyone is told Graph cannot reach.
 *
 *  It can: they are items in the Notes folder wearing a particular message
 *  class, so the mail permission already covers them. The window reads them,
 *  makes them, throws them away, and does the thing Outlook cannot, which is
 *  to put one in the vault as a real note. */
class NotesModal extends Modal {
	private filter = "";
	private problem = "";
	private openId: string | null = null;
	private openBody = "";
	private listEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Outlook notes");
		makeMovable(this.app, this, "powerdesk:notes-window", { w: 640, h: 680 });
		const c = this.contentEl;
		c.addClass("pcal-notes");

		const add = c.createDiv("pcal-tasks-add");
		const input = add.createEl("input", { attr: { type: "text", placeholder: "Write a note and press Enter" } });
		const submit = () => {
			const text = input.value.trim();
			if (!text) return;
			input.value = "";
			void this.plugin.addStickyNote(text).then(() => this.draw());
		};
		input.addEventListener("keydown", (e) => {
			if (e.key === "Enter") submit();
		});
		add.createEl("button", { text: "Add", cls: "mod-cta" }).addEventListener("click", submit);

		const search = c.createEl("input", { cls: "pcal-people-search", attr: { type: "search", placeholder: "Search notes..." } });
		search.addEventListener("input", () => {
			this.filter = search.value.toLowerCase();
			this.draw();
		});

		this.listEl = c.createDiv("pcal-notes-list");
		this.listEl.createDiv({ cls: "pcal-when-note", text: "Reading your notes..." });
		void this.plugin.loadStickyNotes().then((r) => {
			this.problem = r.ok ? "" : r.reason;
			this.draw();
		});

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private draw() {
		const host = this.listEl;
		host.empty();
		if (this.problem) {
			host.createDiv({ cls: "pcal-mail-error", text: this.problem });
			return;
		}
		const all = this.plugin.stickyNotes();
		const notes = this.filter ? all.filter((n) => `${n.title} ${n.preview}`.toLowerCase().includes(this.filter)) : all;
		if (!all.length) {
			host.createDiv({ cls: "pcal-when-note", text: "No notes in this mailbox yet. Write one above and it appears in Outlook too." });
			return;
		}
		if (!notes.length) {
			host.createDiv({ cls: "pcal-when-note", text: "Nothing matches that." });
			return;
		}
		for (const n of notes) {
			const row = host.createDiv("pcal-notes-row");
			row.toggleClass("is-open", this.openId === n.id);
			const mid = row.createDiv("pcal-notes-mid");
			mid.createDiv({ cls: "pcal-notes-title", text: n.title });
			mid.createDiv({ cls: "pcal-notes-sub", text: `${fmtDayShort(keyOfMs(n.changedMs || Date.now()), true)}${n.preview ? ` · ${n.preview.slice(0, 90)}` : ""}` });
			if (this.openId === n.id) mid.createDiv({ cls: "pcal-notes-body", text: this.openBody || "..." });
			const acts = row.createDiv("pcal-people-acts");
			const act = (icon: string, label: string, run: () => void) => {
				const b = acts.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
				setIcon(b, icon);
				b.addEventListener("click", (e) => {
					e.stopPropagation();
					run();
				});
			};
			act("file-plus", "Save to the vault", () => void this.plugin.stickyNoteToVault(n).then(() => this.close()));
			act("trash-2", "Delete the note", () => {
				new ConfirmModal(this.app, `Delete "${n.title}"?`, "A note is an item in your mailbox, so this puts it in Deleted Items where Outlook can still recover it.", "Delete", () => {
					void this.plugin.removeStickyNote(n.id).then(() => this.draw());
				}).open();
			});
			row.addEventListener("click", () => {
				if (this.openId === n.id) {
					this.openId = null;
					this.draw();
					return;
				}
				this.openId = n.id;
				this.openBody = "";
				this.draw();
				void this.plugin.stickyNoteBody(n.id).then((b) => {
					this.openBody = stripHtml(b);
					this.draw();
				});
			});
		}
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Microsoft To Do, and the flagged mail beside it.
 *
 *  Two lists in one window because that is what a task list actually is for
 *  most people: the things they wrote down, and the messages they flagged
 *  meaning "come back to this". Outlook shows both in its To-Do bar for the
 *  same reason. The flagged half needs no permission at all, so it works
 *  even before the tasks half has been granted one. */
class TasksModal extends Modal {
	private listId = "";
	private showDone = false;
	private bodyEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onOpenMail: (m: PCMail) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Tasks");
		makeMovable(this.app, this, "powerdesk:tasks-window", { w: 640, h: 700 });
		this.contentEl.addClass("pcal-tasks");
		this.bodyEl = this.contentEl.createDiv();
		this.draw();
		void this.plugin.loadTaskLists().then(async (lists) => {
			const pick = lists.find((l) => l.isDefault) ?? lists[0];
			if (pick) {
				this.listId = pick.id;
				await this.plugin.loadTasks(pick.id);
			}
			this.draw();
		});
	}

	private draw() {
		const c = this.bodyEl;
		c.empty();
		const lists = this.plugin.todoListsFor();

		if (this.plugin.tasksNeedReconnect())
			c.createDiv({ cls: "pcal-mail-error", text: "Reconnect your accounts in settings to see your Microsoft To Do lists here. Flagged mail below needs no permission and works either way." });

		if (lists.length) {
			const head = c.createDiv("pcal-tasks-head");
			const sel = head.createEl("select", { cls: "dropdown" });
			for (const l of lists) sel.createEl("option", { value: l.id, text: l.name });
			sel.value = this.listId;
			sel.addEventListener("change", () => {
				this.listId = sel.value;
				void this.plugin.loadTasks(this.listId).then(() => this.draw());
			});
			const doneBtn = head.createEl("button", { text: this.showDone ? "Hide done" : "Show done" });
			doneBtn.addEventListener("click", () => {
				this.showDone = !this.showDone;
				this.draw();
			});

			// adding one is a line and a return, not a dialog
			const add = c.createDiv("pcal-tasks-add");
			const input = add.createEl("input", { attr: { type: "text", placeholder: "Add a task and press Enter" } });
			const due = add.createEl("input", { attr: { type: "date" } });
			const submit = () => {
				const title = input.value.trim();
				if (!title) return;
				const dueMs = /^\d{4}-\d{2}-\d{2}$/.test(due.value) ? msOfKey(due.value) : null;
				input.value = "";
				void this.plugin.addTask(this.listId, title, dueMs).then(() => this.draw());
			};
			input.addEventListener("keydown", (e) => {
				if (e.key === "Enter") submit();
			});
			add.createEl("button", { text: "Add", cls: "mod-cta" }).addEventListener("click", submit);

			const tasks = this.plugin.tasksIn(this.listId).filter((t) => this.showDone || !t.done);
			const box = c.createDiv("pcal-tasks-list");
			if (!tasks.length) box.createDiv({ cls: "pcal-when-note", text: this.showDone ? "Nothing in this list." : "Nothing left in this list." });
			// overdue first, then by due date, then the undated
			const today = msOfKey(keyOfDate(new Date()));
			for (const t of [...tasks].sort((a, b) => (a.dueMs ?? Infinity) - (b.dueMs ?? Infinity))) {
				const row = box.createDiv("pcal-tasks-row");
				row.toggleClass("is-done", t.done);
				const tick = row.createSpan("pcal-tasks-tick");
				setIcon(tick, t.done ? "check-square" : "square");
				tick.addEventListener("click", () => void this.plugin.setTaskDone(t, !t.done).then(() => this.draw()));
				const mid = row.createDiv("pcal-tasks-mid");
				mid.createDiv({ cls: "pcal-tasks-title", text: t.title });
				if (t.dueMs != null) {
					const late = !t.done && t.dueMs < today;
					const when = mid.createDiv({ cls: "pcal-tasks-due", text: `${late ? "Overdue, " : "Due "}${fmtDayShort(keyOfMs(t.dueMs), true)}` });
					when.toggleClass("is-late", late);
				}
				if (t.importance === "high") row.createSpan({ cls: "pcal-mail-bang", text: "!" });
				const del = row.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Delete task" } });
				setIcon(del, "trash-2");
				del.addEventListener("click", () => void this.plugin.removeTask(t).then(() => this.draw()));
			}
		}

		const flagged = this.plugin.flaggedMail();
		c.createDiv({ cls: "pcal-shortcuts-head", text: `Flagged mail${flagged.length ? ` (${flagged.length})` : ""}` });
		const fbox = c.createDiv("pcal-tasks-list");
		if (!flagged.length) fbox.createDiv({ cls: "pcal-when-note", text: "Nothing flagged. Press S on a message to flag it." });
		for (const m of flagged) {
			const row = fbox.createDiv("pcal-tasks-row");
			const tick = row.createSpan("pcal-tasks-tick");
			setIcon(tick, "flag");
			tick.addClass("is-flag");
			tick.setAttribute("aria-label", "Clear the flag");
			tick.addEventListener("click", (e) => {
				e.stopPropagation();
				void this.plugin.setMailFlag(m, false);
				this.draw();
			});
			const mid = row.createDiv("pcal-tasks-mid");
			mid.createDiv({ cls: "pcal-tasks-title", text: m.subject });
			mid.createDiv({ cls: "pcal-tasks-due", text: `${m.from} · ${fmtDayShort(keyOfMs(m.receivedMs), true)}` });
			row.addEventListener("click", () => {
				this.onOpenMail(m);
				this.close();
			});
		}

		const btns = this.contentEl.querySelector(".pcal-tasks-btns") ?? this.contentEl.createDiv("pcal-modal-btns pcal-compose-btns pcal-tasks-btns");
		btns.empty();
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** The people this vault already knows about, and what can be done with one.
 *
 *  Built from correspondence rather than from an address book: who you have
 *  written to, who has written to you, and who you have met. That is a
 *  different list from Outlook's Contacts and usually a more useful one,
 *  because it is ordered by how much you actually deal with someone rather
 *  than by whether anyone remembered to save them. The actions are the ones
 *  only this plugin can offer, which is the point of it living here: their
 *  mail, their meetings, and their page in the vault, all from one row. */
class PeopleModal extends Modal {
	private filter = "";
	private listEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onSearchMail: (q: string) => void,
		private onSearchEvents: (name: string) => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("People");
		makeMovable(this.app, this, "powerdesk:people-window", { w: 640, h: 680 });
		const c = this.contentEl;
		c.addClass("pcal-people");
		void this.plugin.ensureSentContacts();
		void this.plugin.ensureSavedContacts().then(() => this.draw());
		if (this.plugin.contactsNeedReconnect())
			c.createDiv({
				cls: "pcal-mail-error",
				text: "Reconnect your accounts in settings to include the contacts saved in your mailbox. Until then this lists everyone you correspond with, which needs no permission.",
			});

		const search = c.createEl("input", { cls: "pcal-people-search", attr: { type: "search", placeholder: "Search people..." } });
		search.addEventListener("input", () => {
			this.filter = search.value;
			this.draw();
		});
		this.listEl = c.createDiv("pcal-people-list");
		this.draw();

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
		window.setTimeout(() => search.focus(), 20);
	}

	private draw() {
		const host = this.listEl;
		host.empty();
		const index = this.plugin.people();
		const hits = matchContacts(index, this.filter, 300) as PersonCard[];
		if (!index.length) {
			host.createDiv({ cls: "pcal-when-note", text: "Nobody yet. Open the inbox once so there is some correspondence to read." });
			return;
		}
		if (!hits.length) {
			host.createDiv({ cls: "pcal-when-note", text: "Nobody matches that." });
			return;
		}
		host.createDiv({ cls: "pcal-when-note", text: `${hits.length} of ${index.length}` });
		for (const p of hits) {
			const row = host.createDiv("pcal-people-row");
			const av = row.createDiv("pcal-mail-avatar");
			const photo = this.plugin.settings.mailPhotos ? this.plugin.photoFor(p.email) : null;
			if (photo) {
				av.addClass("has-photo");
				av.style.backgroundImage = `url("${photo}")`;
			} else {
				av.style.backgroundColor = avatarColor(p.name || p.email);
				av.setText(avatarInitials(p.name || p.email));
			}
			const mid = row.createDiv("pcal-people-mid");
			const nameRow = mid.createDiv("pcal-people-name");
			nameRow.createSpan({ text: p.name || p.email });
			// a saved contact is marked, so it is clear which names came from
			// the address book and which were inferred from correspondence
			if (p.saved) nameRow.createSpan({ cls: "pcal-people-saved", text: "contact" });
			const bits = [p.email];
			if (p.title || p.company) bits.push([p.title, p.company].filter(Boolean).join(", "));
			if (p.phone) bits.push(p.phone);
			bits.push(p.count ? `${p.count} message${p.count === 1 ? "" : "s"}, last ${fmtDayShort(keyOfMs(p.lastMs), true)}` : "no mail either way");
			mid.createDiv({ cls: "pcal-people-sub", text: bits.join(" · ") });
			const acts = row.createDiv("pcal-people-acts");
			const act = (icon: string, label: string, run: () => void) => {
				const b = acts.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": label } });
				setIcon(b, icon);
				b.addEventListener("click", (e) => {
					e.stopPropagation();
					run();
				});
			};
			act("pencil", `New mail to ${p.name || p.email}`, () => new RichComposeModal(this.app, this.plugin, { mode: "new", to: p.email }).open());
			act("inbox", "Their mail", () => {
				this.onSearchMail(`from:${p.email}`);
				this.close();
			});
			act("calendar-days", "Meetings with them", () => {
				this.onSearchEvents(p.name || p.email);
				this.close();
			});
			act("file-text", "Their page in the vault", () => {
				void this.plugin.openPersonPage(p.name || p.email);
				this.close();
			});
			row.addEventListener("click", () => void navigator.clipboard.writeText(p.email).then(() => new Notice(`Power Desk: copied ${p.email}.`)));
		}
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** The calendar's own Find: fields at the top, results underneath, and a
 *  click to land on the day.
 *
 *  Results live in the window rather than replacing the calendar, because a
 *  search of a year is a list you scan while the calendar stays where it
 *  was, and because the answer to "when did I meet Deanna about Kore" is
 *  usually read rather than navigated to. */
class EventSearchModal extends Modal {
	private q: EventQuery = { words: "", title: "", people: "", location: "", calendar: "", onlineOnly: false, allDayOnly: false, withPeopleOnly: false };
	private fromKey: string;
	private toKey: string;
	private resultsEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onPick: (key: string) => void,
		startPerson?: string
	) {
		super(app);
		if (startPerson) this.q.people = startPerson;
		const today = keyOfDate(new Date());
		// a year back and a year on: far enough for "when was that", small
		// enough that it does not fetch a decade to answer one question
		this.fromKey = addDays(today, -365);
		this.toKey = addDays(today, 365);
	}

	onOpen() {
		this.titleEl.setText("Find events");
		makeMovable(this.app, this, "powerdesk:event-search-window", { w: 680, h: 700 });
		const c = this.contentEl;
		c.addClass("pcal-search-modal");

		const text = (name: string, desc: string, key: "words" | "title" | "people" | "location", placeholder: string) =>
			new Setting(c)
				.setName(name)
				.setDesc(desc)
				.addText((t) =>
					t
						.setPlaceholder(placeholder)
						.setValue(this.q[key])
						.onChange((v) => {
							this.q[key] = v;
							this.run();
						})
				);
		text("Words", "Anywhere in the event, including its description.", "words", "kore tank");
		text("Title", "Every word must be in the title.", "title", "review");
		text("With", "An organizer or attendee.", "people", "deanna");
		text("Location", "Where it is, or the room.", "location", "room 4");

		const sources = this.plugin.sources();
		if (sources.length > 1) {
			new Setting(c).setName("Calendar").addDropdown((d) => {
				d.addOption("", "Every calendar");
				for (const s of sources) d.addOption(s.key, s.label);
				d.setValue(this.q.calendar).onChange((v) => {
					this.q.calendar = v;
					this.run();
				});
			});
		}
		const toggle = (name: string, key: "onlineOnly" | "allDayOnly" | "withPeopleOnly") =>
			new Setting(c).setName(name).addToggle((t) =>
				t.setValue(this.q[key]).onChange((v) => {
					this.q[key] = v;
					this.run();
				})
			);
		toggle("Has a join link", "onlineOnly");
		toggle("All-day only", "allDayOnly");
		toggle("Has other people", "withPeopleOnly");

		const range = (name: string, get: () => string, set: (v: string) => void) =>
			new Setting(c).setName(name).addText((t) => {
				t.inputEl.type = "date";
				t.setValue(get()).onChange((v) => {
					if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
						set(v);
						this.run();
					}
				});
			});
		range("From", () => this.fromKey, (v) => (this.fromKey = v));
		range("To", () => this.toKey, (v) => (this.toKey = v));

		c.createDiv({ cls: "pcal-shortcuts-head", text: "Results" });
		this.resultsEl = c.createDiv("pcal-eventsearch-results");
		this.run();

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private run() {
		const host = this.resultsEl;
		host.empty();
		if (eventQueryIsEmpty(this.q)) {
			host.createDiv({ cls: "pcal-when-note", text: "Fill something in above and the matches appear here." });
			return;
		}
		// the range has to be fetched before it can be searched; what is
		// already cached shows at once and the rest arrives on the next pass
		this.plugin.ensureWindow(this.fromKey, this.toKey, false);
		const hits = matchEvents(this.plugin.eventsForWindow(this.fromKey, this.toKey), this.q);
		if (!hits.length) {
			host.createDiv({ cls: "pcal-when-note", text: "Nothing in this range matches." });
			return;
		}
		const s = this.plugin.settings;
		host.createDiv({ cls: "pcal-when-note", text: `${hits.length} event${hits.length === 1 ? "" : "s"}` });
		for (const ev of hits.slice(0, 300)) {
			const key = eventDaySpan(ev).startKey;
			const row = host.createDiv("pcal-eventsearch-row");
			row.createDiv({ cls: "pcal-eventsearch-when", text: `${fmtDayShort(key, true)}${ev.allDay ? "" : ` ${fmtTimeOfMs(ev.startMs, s.use24h, true)}`}` });
			const mid = row.createDiv("pcal-eventsearch-mid");
			mid.createDiv({ cls: "pcal-eventsearch-title", text: ev.title });
			const bits = [ev.calendarName, ev.location, (ev.attendees ?? []).slice(0, 3).join(", ")].filter(Boolean).join(" · ");
			if (bits) mid.createDiv({ cls: "pcal-eventsearch-sub", text: bits });
			row.addEventListener("click", () => {
				this.onPick(key);
				this.close();
			});
		}
		if (hits.length > 300) host.createDiv({ cls: "pcal-when-note", text: `Showing the first 300. Narrow the range or add a word to see the rest.` });
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Outlook's Advanced Find, in the shape this mailbox actually supports.
 *
 *  The fields compose the same query language the search box takes, and the
 *  composed text is shown back, so the dialog is a way into the syntax
 *  rather than a separate search that happens to look similar. */
class SearchModal extends Modal {
	private f: SearchFields = { words: "", from: "", subject: "", phrase: "", unread: false, flagged: false, attachments: false, after: "", before: "" };
	private previewEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private view: MailView,
		start?: string
	) {
		super(app);
		if (start?.trim()) this.f.words = start.trim();
	}

	onOpen() {
		this.titleEl.setText("Search mail");
		makeMovable(this.app, this, "powerdesk:search-window", { w: 620, h: 620 });
		const c = this.contentEl;
		c.addClass("pcal-search-modal");

		const text = (name: string, desc: string, key: "words" | "from" | "subject" | "phrase", placeholder: string) =>
			new Setting(c)
				.setName(name)
				.setDesc(desc)
				.addText((t) =>
					t
						.setPlaceholder(placeholder)
						.setValue(this.f[key])
						.onChange((v) => {
							this.f[key] = v;
							this.paint();
						})
				);
		text("Words", "Anywhere in the message. Each word must appear.", "words", "invoice pulseway");
		text("Exact phrase", "These words, in this order.", "phrase", "tank management");
		text("From", "Name or address; several separated by commas.", "from", "deanna, jira@");
		text("Subject", "Each word must be in the subject.", "subject", "invoice");

		const toggle = (name: string, key: "unread" | "flagged" | "attachments") =>
			new Setting(c).setName(name).addToggle((t) =>
				t.setValue(this.f[key]).onChange((v) => {
					this.f[key] = v;
					this.paint();
				})
			);
		toggle("Unread only", "unread");
		toggle("Flagged only", "flagged");
		toggle("Has an attachment", "attachments");

		const date = (name: string, key: "after" | "before") =>
			new Setting(c).setName(name).addText((t) => {
				t.inputEl.type = "date";
				t.setValue(this.f[key]).onChange((v) => {
					this.f[key] = v;
					this.paint();
				});
			});
		date("On or after", "after");
		date("On or before", "before");

		c.createDiv({ cls: "pcal-shortcuts-head", text: "This searches for" });
		this.previewEl = c.createDiv("pcal-search-preview");
		this.paint();

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		btns.createEl("button", { text: "Search the mailbox", cls: "mod-cta pcal-send-btn" }).addEventListener("click", () => this.run(true));
		btns.createEl("button", { text: "Search this device" }).addEventListener("click", () => this.run(false));
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Close" }).addEventListener("click", () => this.close());
	}

	private query(): string {
		return buildQuery(this.f);
	}

	private paint() {
		const q = this.query();
		this.previewEl.setText(q || "everything (nothing entered yet)");
		this.previewEl.toggleClass("is-empty", !q);
	}

	private run(everywhere: boolean) {
		const q = this.query();
		if (!q) {
			new Notice("Power Desk: fill in something to search for.");
			return;
		}
		this.view.runSearch(q, everywhere);
		this.close();
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** Pick one folder in one account, for a rule's destination. */
class FolderPickModal extends FuzzySuggestModal<{ folderId: string; name: string; path: string }> {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private accountId: string,
		private onPick: (folderId: string, name: string) => void
	) {
		super(app);
		this.setPlaceholder("Move matching mail to...");
	}

	getItems() {
		return this.plugin.moveTargets([this.accountId]);
	}

	getItemText(c: { path: string }): string {
		return c.path;
	}

	onChooseItem(c: { folderId: string; name: string }): void {
		this.onPick(c.folderId, c.name);
	}
}

/** The rules an account's mailbox runs, and the way in to editing them. */
class RulesModal extends Modal {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app);
	}

	private loading = true;

	onOpen() {
		this.titleEl.setText("Inbox rules");
		makeMovable(this.app, this, "powerdesk:rules-window", { w: 760, h: 600 });
		this.draw();
		void (async () => {
			for (const a of this.plugin.mailAccounts()) await this.plugin.loadRules(a.id);
			this.loading = false;
			this.draw();
		})();
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-rules");
		c.createDiv({
			cls: "pcal-when-note",
			text: "Rules run in your mailbox, not in Obsidian, so they file mail whether or not this is open and on every device you read mail from. They are the same rules Outlook shows.",
		});
		for (const a of this.plugin.mailAccounts()) {
			c.createDiv({ cls: "pcal-shortcuts-head", text: this.plugin.nameOf(a) });
			const rules = this.plugin.rulesFor(a.id);
			const err = this.plugin.rulesErrorFor(a.id);
			if (err) c.createDiv({ cls: "pcal-mail-error", text: err });
			else if (!rules.length) c.createDiv({ cls: "pcal-when-note", text: this.loading ? "Reading the rules from your mailbox..." : "No rules yet." });
			for (const r of rules) {
				const row = c.createDiv("pcal-rule-row");
				row.toggleClass("is-off", !r.isEnabled);
				const tx = row.createDiv("pcal-rule-text");
				tx.createDiv({ cls: "pcal-rule-name", text: r.displayName });
				const edit = ruleToEdit(r);
				const dest = edit.moveToFolderId ? this.plugin.moveTargets([a.id]).find((f) => f.folderId === edit.moveToFolderId)?.name : undefined;
				tx.createDiv({ cls: "pcal-rule-what", text: `${r.isEnabled ? "" : "Off. "}${ruleSummary(edit, dest)}` });
				if (ruleHasUnknownParts(r)) tx.createDiv({ cls: "pcal-rule-warn", text: "This rule also does things Power Desk does not show. Editing it here keeps them." });
				row.createEl("button", { text: r.isEnabled ? "Turn off" : "Turn on" }).addEventListener("click", () => {
					void this.plugin.saveRule(a.id, { ...edit, enabled: !r.isEnabled }, r).then(() => this.draw());
				});
				row.createEl("button", { text: "Edit" }).addEventListener("click", () => new RuleEditModal(this.app, this.plugin, a.id, r, null, () => this.draw()).open());
				row.createEl("button", { text: "Delete" }).addEventListener("click", () => {
					new ConfirmModal(this.app, `Delete "${r.displayName}"?`, "The rule stops running in your mailbox. Mail it already filed stays where it is.", "Delete rule", () => {
						void this.plugin.removeRule(a.id, r.id).then(() => this.draw());
					}).open();
				});
			}
			// creating would fail the same way reading did, so do not offer it
			if (err) continue;
			const add = c.createDiv("pcal-modal-btns");
			add.createEl("button", { text: "New rule", cls: "mod-cta" }).addEventListener("click", () => new RuleEditModal(this.app, this.plugin, a.id, null, null, () => this.draw()).open());
		}
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** The rule editor: the conditions this plugin offers, the actions it
 *  offers, and a live sentence saying what the whole thing will do. */
class RuleEditModal extends Modal {
	private edit: RuleEdit;
	private destName = "";
	private summaryEl!: HTMLElement;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private accountId: string,
		private existing: GraphRule | null,
		prefill: Partial<RuleEdit> | null,
		private onSaved: () => void
	) {
		super(app);
		this.edit = existing ? ruleToEdit(existing) : { ...EMPTY_RULE, ...(prefill ?? {}) };
		if (this.edit.moveToFolderId) this.destName = this.plugin.moveTargets([accountId]).find((f) => f.folderId === this.edit.moveToFolderId)?.name ?? "";
	}

	onOpen() {
		this.titleEl.setText(this.existing ? "Edit rule" : "New rule");
		// its own key: the editor is a tall form and the list beside it is
		// not, so they should not have to agree on a size
		makeMovable(this.app, this, "powerdesk:rule-edit-window", { w: 660, h: 680 });
		const c = this.contentEl;
		c.addClass("pcal-rule-edit");

		new Setting(c).setName("Name").addText((t) =>
			t.setValue(this.edit.name).onChange((v) => {
				this.edit.name = v;
			})
		);

		c.createDiv({ cls: "pcal-shortcuts-head", text: "When a message arrives and" });
		const text = (label: string, key: "fromContains" | "subjectContains" | "bodyContains" | "toContains", placeholder: string) =>
			new Setting(c)
				.setName(label)
				.addText((t) =>
					t
						.setPlaceholder(placeholder)
						.setValue(this.edit[key])
						.onChange((v) => {
							this.edit[key] = v;
							this.paintSummary();
						})
				)
				.setDesc("Comma separated; any one of them matches.");
		text("From contains", "fromContains", "jira@, no-reply@");
		text("Subject contains", "subjectContains", "invoice");
		text("Body contains", "bodyContains", "");
		text("Sent to contains", "toContains", "");
		new Setting(c).setName("Has an attachment").addToggle((t) =>
			t.setValue(this.edit.hasAttachments).onChange((v) => {
				this.edit.hasAttachments = v;
				this.paintSummary();
			})
		);
		new Setting(c).setName("Marked high importance").addToggle((t) =>
			t.setValue(this.edit.highImportance).onChange((v) => {
				this.edit.highImportance = v;
				this.paintSummary();
			})
		);

		c.createDiv({ cls: "pcal-shortcuts-head", text: "Then" });
		const moveSt = new Setting(c).setName("Move to folder").setDesc(this.destName || "Not set");
		moveSt.addButton((b) =>
			b.setButtonText("Choose...").onClick(() =>
				new FolderPickModal(this.app, this.plugin, this.accountId, (id, name) => {
					this.edit.moveToFolderId = id;
					this.destName = name;
					moveSt.setDesc(name);
					this.paintSummary();
				}).open()
			)
		);
		moveSt.addExtraButton((b) =>
			b
				.setIcon("x")
				.setTooltip("Do not move")
				.onClick(() => {
					this.edit.moveToFolderId = "";
					this.destName = "";
					moveSt.setDesc("Not set");
					this.paintSummary();
				})
		);
		new Setting(c).setName("Mark as read").addToggle((t) =>
			t.setValue(this.edit.markAsRead).onChange((v) => {
				this.edit.markAsRead = v;
				this.paintSummary();
			})
		);
		new Setting(c).setName("Set importance").addDropdown((d) => {
			d.addOption("", "Leave alone");
			d.addOption("low", "Low");
			d.addOption("normal", "Normal");
			d.addOption("high", "High");
			d.setValue(this.edit.markImportance).onChange((v) => {
				this.edit.markImportance = v as RuleEdit["markImportance"];
				this.paintSummary();
			});
		});
		new Setting(c)
			.setName("Delete it")
			.setDesc("Into Deleted Items, where it can still be recovered.")
			.addToggle((t) =>
				t.setValue(this.edit.deleteIt).onChange((v) => {
					this.edit.deleteIt = v;
					this.paintSummary();
				})
			);
		new Setting(c)
			.setName("Stop processing more rules")
			.setDesc("Later rules do not see a message this one matched.")
			.addToggle((t) =>
				t.setValue(this.edit.stopProcessing).onChange((v) => {
					this.edit.stopProcessing = v;
					this.paintSummary();
				})
			);

		this.summaryEl = c.createDiv("pcal-rule-summary");
		this.paintSummary();
		if (this.existing && ruleHasUnknownParts(this.existing)) {
			c.createDiv({
				cls: "pcal-when-note",
				text: "This rule also carries conditions or actions Power Desk does not show, set in Outlook. They are kept exactly as they are when you save here.",
			});
		}

		const btns = c.createDiv("pcal-modal-btns");
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
		btns.createEl("button", { text: "Save", cls: "mod-cta" }).addEventListener("click", () => {
			void this.plugin.saveRule(this.accountId, this.edit, this.existing ?? undefined).then((ok) => {
				if (!ok) return;
				this.close();
				this.onSaved();
			});
		});
	}

	private paintSummary() {
		this.summaryEl?.setText(ruleSummary(this.edit, this.destName));
	}

	onClose() {
		this.contentEl.empty();
	}
}

interface MoveChoice {
	accountId: string;
	accountLabel: string;
	folderId: string;
	name: string;
	path: string;
	recent: boolean;
}

/** Where to file: a fuzzy list over the whole folder tree, the handful you
 *  actually use at the top.
 *
 *  A picker rather than a submenu because the tree here is deep and named
 *  after people and projects; typing three letters beats walking it. Graph
 *  moves within one mailbox, so a selection spanning two accounts files the
 *  part that belongs to the chosen folder's account and says how many that
 *  was. */
class MoveToFolderModal extends FuzzySuggestModal<MoveChoice> {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private targets: PCMail[],
		private onDone: () => void
	) {
		super(app);
		const n = targets.length;
		this.setPlaceholder(n > 1 ? `Move ${n} messages to...` : "Move to...");
		this.setInstructions([
			{ command: "↑↓", purpose: "navigate" },
			{ command: "↵", purpose: "move" },
			{ command: "esc", purpose: "cancel" },
		]);
	}

	getItems(): MoveChoice[] {
		const accountIds = [...new Set(this.targets.map((t) => t.accountId))];
		const all = this.plugin.moveTargets(accountIds);
		const recentKeys = this.plugin.settings.mailRecentFolders.map((r) => `${r.accountId}:${r.folderId}`);
		const key = (c: { accountId: string; folderId: string }) => `${c.accountId}:${c.folderId}`;
		const recent = recentKeys.map((k) => all.find((c) => key(c) === k)).filter((c): c is (typeof all)[number] => !!c);
		const rest = all.filter((c) => !recentKeys.includes(key(c)));
		return [...recent.map((c) => ({ ...c, recent: true })), ...rest.map((c) => ({ ...c, recent: false }))];
	}

	getItemText(c: MoveChoice): string {
		return `${c.path} ${c.accountLabel}`;
	}

	renderSuggestion(m: { item: MoveChoice }, el: HTMLElement): void {
		el.addClass("pcal-palette-row");
		el.createSpan({ cls: "pcal-palette-label", text: m.item.path });
		el.createSpan({ cls: "pcal-palette-hint", text: m.item.recent ? `${m.item.accountLabel} — recent` : m.item.accountLabel });
	}

	onChooseItem(c: MoveChoice): void {
		const mine = this.targets.filter((t) => t.accountId === c.accountId);
		const skipped = this.targets.length - mine.length;
		void this.plugin.moveMail(this.targets, c.accountId, c.folderId, c.name).then((n) => {
			if (n) new Notice(n > 1 ? `Power Desk: moved ${n} messages to ${c.name}.` : `Power Desk: moved to ${c.name}.`);
			if (skipped) new Notice(`Power Desk: ${skipped} message${skipped === 1 ? "" : "s"} stayed put, being in another account. A move cannot cross mailboxes.`, 8000);
			this.onDone();
		});
	}
}

/** What is snoozed and when it comes back, with a way to pull one forward.
 *  A snooze you cannot see is a message you have lost, so this exists even
 *  though the Snoozed folder itself is browsable in the folder tree. */
class SnoozedModal extends Modal {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private onChange: () => void
	) {
		super(app);
	}

	onOpen() {
		this.titleEl.setText("Snoozed mail");
		this.draw();
	}

	private draw() {
		const c = this.contentEl;
		c.empty();
		c.addClass("pcal-snoozed");
		const rows = this.plugin.snoozedFor();
		if (!rows.length) {
			c.createDiv({ cls: "pcal-when-note", text: "Nothing is snoozed." });
			return;
		}
		for (const rec of rows) {
			const row = c.createDiv("pcal-snoozed-row");
			const txt = row.createDiv("pcal-snoozed-text");
			txt.createDiv({ cls: "pcal-snoozed-subject", text: rec.subject });
			txt.createDiv({ cls: "pcal-snoozed-when", text: `Back ${fmtWhen(rec.dueMs, this.plugin.settings.use24h)}` });
			row.createEl("button", { text: "Bring back now" }).addEventListener("click", () => {
				void this.plugin.unsnooze(rec).then(() => {
					this.draw();
					this.onChange();
				});
			});
		}
		c.createDiv({
			cls: "pcal-when-note",
			text: "Snoozed mail waits in the Snoozed folder in your mailbox. It returns to the inbox when its time comes and Obsidian is running; if it is closed, the message comes back the next time you open it.",
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

/** One thing the mail palette can do. */
interface PaletteItem {
	/** What the row reads as, and what the fuzzy match runs against. */
	label: string;
	/** The muted right-hand column: the account, the folder, a state. */
	hint?: string;
	/** Extra words the match should find it by but the row need not show. */
	terms?: string;
	run: () => void;
}

/** The mail command palette: every folder in every account and every action
 *  that applies right now, behind one fuzzy search.
 *
 *  Superhuman's Cmd-K is the feature people quote about it, and the reason it
 *  works is that jumping and doing live in the same box: you never have to
 *  know whether the thing you want is a place or a verb. Obsidian already
 *  ships the fuzzy modal, so this is mostly the business of assembling a good
 *  list. Actions that need a message are only listed when one is open. */
class MailPaletteModal extends FuzzySuggestModal<PaletteItem> {
	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private view: MailView
	) {
		super(app);
		this.setPlaceholder("Jump to a folder, or run a command...");
		this.setInstructions([
			{ command: "↑↓", purpose: "navigate" },
			{ command: "↵", purpose: "run" },
			{ command: "esc", purpose: "dismiss" },
		]);
	}

	getItems(): PaletteItem[] {
		return this.view.paletteItems();
	}

	getItemText(i: PaletteItem): string {
		return `${i.label} ${i.hint ?? ""} ${i.terms ?? ""}`;
	}

	renderSuggestion(m: { item: PaletteItem }, el: HTMLElement): void {
		el.addClass("pcal-palette-row");
		el.createSpan({ cls: "pcal-palette-label", text: m.item.label });
		if (m.item.hint) el.createSpan({ cls: "pcal-palette-hint", text: m.item.hint });
	}

	onChooseItem(i: PaletteItem): void {
		i.run();
	}
}

/** The shortcut card, on ? like every keyboard-first client has one. */
class MailShortcutsModal extends Modal {
	onOpen() {
		this.titleEl.setText("Mail shortcuts");
		const c = this.contentEl;
		c.addClass("pcal-shortcuts");
		const groups: [string, [string, string][]][] = [
			[
				"Moving",
				[
					["J  /  ↓", "Next message"],
					["K  /  ↑", "Previous message"],
					["↵", "Open the message"],
					["U", "Back to the list"],
					["→  /  ←", "Open or close the conversation"],
					["O", "Toggle the conversation"],
				],
			],
			[
				"Acting",
				[
					["R", "Reply"],
					["A", "Reply all"],
					["F", "Forward"],
					["C", "New mail"],
					["E", "Archive"],
					["X", "Tick the message"],
					["S", "Flag or clear the flag"],
					["G", "Categorize"],
					["B", "Snooze until later"],
					["V", "Move to a folder"],
					["Shift + I", "Mark read"],
					["Shift + U", "Mark unread"],
					["Del", "Delete"],
					["Shift + Del", "Delete permanently"],
				],
			],
			[
				"Finding",
				[
					["/", "Search as you type"],
					["↵ in search", "Widen it to the whole mailbox"],
					["Ctrl / Cmd + K", "Command palette"],
					["?", "This card"],
					["Esc", "Clear the search or the selection"],
				],
			],
		];
		for (const [title, rows] of groups) {
			c.createDiv({ cls: "pcal-shortcuts-head", text: title });
			for (const [k, what] of rows) {
				const row = c.createDiv("pcal-shortcuts-row");
				row.createSpan({ cls: "pcal-shortcuts-key", text: k });
				row.createSpan({ cls: "pcal-shortcuts-what", text: what });
			}
		}
		c.createDiv({ cls: "pcal-shortcuts-head", text: "Search terms" });
		for (const [k, what] of [
			["from:name", "The sender's name or address"],
			["subject:word", "The subject only"],
			["is:unread", "Unread, or is:read"],
			["is:flagged", "Flagged for follow-up"],
			["has:attachment", "Carries a file"],
			['"exact words"', "That run, in that order"],
		] as [string, string][]) {
			const row = c.createDiv("pcal-shortcuts-row");
			row.createSpan({ cls: "pcal-shortcuts-key", text: k });
			row.createSpan({ cls: "pcal-shortcuts-what", text: what });
		}
		c.createDiv({
			cls: "pcal-shortcuts-note",
			text: "Shortcuts work while the message list has focus, never while you are typing in a field. Typing in the search box searches the mail already on this device, which is instant; Enter runs the same words against the whole mailbox.",
		});
	}

	onClose() {
		this.contentEl.empty();
	}
}

/* Separators for the list signature. Control characters, because the fields
 * being joined are subjects and sender names, which can hold every printable
 * character there is: without these, "AB" then "C" would sign identically to
 * "A" then "BC" and a row would quietly stop redrawing. */
const SIG_FIELD = "\u001f";
const SIG_ROW = "\u001e";

/** One thing that can appear in the message column: an error line, the empty
 *  line, a section band, or a row. `sig` is everything it would draw, so an
 *  unchanged one is left in place; `make` builds it when it is not. */
interface ListItem {
	key: string;
	sig: string;
	make: () => HTMLElement;
}

/** One visible row of the message list, with what it acts on. */
interface MailRow {
	m: PCMail;
	targets: PCMail[];
	thread: MailThread | null;
	child: boolean;
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
	private foldToggleBtn!: HTMLElement;
	private calToolsEl!: HTMLElement;
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
				if (t && (t.instanceOf(HTMLInputElement) || t.instanceOf(HTMLTextAreaElement) || t.isContentEditable)) return true;
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
		this.scope.register([], "Escape", () => {
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
		// first in the row and the same glyph as the mail view's, because it is
		// the same button doing the same thing to the same sidebar: a control
		// that moves depending on which tab you are on is a different control
		this.foldToggleBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Full width: fold the vault's notes away while this tab is open" } });
		setIcon(this.foldToggleBtn, "panel-left");
		this.foldToggleBtn.toggleClass("is-active", this.plugin.focusOn());
		this.foldToggleBtn.addEventListener("click", () => {
			this.plugin.toggleFocus();
			this.syncFoldToggle();
		});
		this.register(this.plugin.watchFocus(() => this.syncFoldToggle()));
		if (!Platform.isPhone) {
			// the notes toggle keeps the panel glyph, so this one takes the
			// columns shape: two panes side by side, which is what it does
			const sbBtn = header.createEl("button", { cls: "pcal-icon-btn", attr: { "aria-label": "Show or hide this calendar's month picker and agenda rail (S)" } });
			setIcon(sbBtn, pickIcon("columns-2", "columns", "layout-list", "list"));
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
		this.calToolsEl = right.createDiv("pcal-cal-tools");
		this.renderCalToolbar();
		const more = right.createEl("button", { cls: "pcal-icon-btn pcal-mail-tool-more", attr: { "aria-label": "Customize the toolbar" } });
		setIcon(more, "settings-2");
		more.addEventListener("click", () =>
			new ToolbarModal(
				this.app,
				this.plugin,
				this.calActions(),
				{
					get: () => this.plugin.settings.calendarToolbar,
					set: (v) => (this.plugin.settings.calendarToolbar = v),
					fallback: DEFAULT_SETTINGS.calendarToolbar,
					leads: "New event",
				},
				() => this.renderCalToolbar()
			).open()
		);
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
								const frag = createFragment();
								frag.createSpan("pcal-sb-menu-dot").style.background = hex;
								frag.appendText(SIDEBAR_COLOR_NAMES[idx] ?? hex);
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

	/** What the calendar header can carry. Everything here acts on the
	 *  calendar as a whole, which is why nothing needs a selection and why
	 *  the nav and view switcher are not among them: those are the view. */
	private calActions(): { id: string; label: string; icon: string; run: (e: MouseEvent) => void }[] {
		return [
			{ id: "filter", label: "Filter", icon: "filter", run: (e) => this.openFilterMenu(e) },
			{ id: "refresh", label: "Refresh", icon: "refresh-cw", run: () => this.refresh(true) },
			{ id: "findEvent", label: "Find event", icon: "search", run: () => new EventFindModal(this.app, this.plugin).open() },
			{ id: "searchEvents", label: "Find events", icon: pickIcon("search-check", "filter", "search"), run: () => this.openEventSearch() },
			{ id: "people", label: "People", icon: "users", run: () => this.plugin.openPeople() },
			{ id: "tasks", label: "Tasks", icon: "check-square", run: () => this.plugin.openTasks() },
			{ id: "notes", label: "Notes", icon: pickIcon("sticky-note", "file-text"), run: () => new NotesModal(this.app, this.plugin).open() },
			{ id: "journal", label: "Journal", icon: pickIcon("book-open", "book", "file-text"), run: () => new JournalModal(this.app, this.plugin).open() },
			{ id: "freeSlots", label: "Copy free slots", icon: "copy", run: () => void this.plugin.copyFreeSlots() },
			{ id: "today", label: "Today", icon: "calendar-check", run: () => this.goToday() },
			{ id: "newMail", label: "New mail", icon: "pencil", run: () => new RichComposeModal(this.app, this.plugin, { mode: "new" }).open() },
			{ id: "inbox", label: "Inbox", icon: "inbox", run: () => void this.plugin.openMailView() },
			// the month has two sensible shapes on paper; the print window is
			// where that is chosen now, with the page in front of you
			{ id: "print", label: "Print", icon: "printer", run: () => this.printCalendar() },
			{ id: "eventsBase", label: "Events base", icon: "database", run: () => void this.plugin.createEventsBase() },
			{ id: "testReminder", label: "Test reminder", icon: "bell", run: () => this.plugin.previewReminder() },
		];
	}

	/** Print what is on screen: the month as its grid, anything else as an
	 *  agenda. A timed grid on paper turns an hour into a few millimetres and
	 *  a busy day into a smear, so the week and day views print as the list
	 *  of what is actually happening, which is what anyone carrying a printed
	 *  day wants from it. */
	printCalendar(style?: "agenda" | "month") {
		if (!Platform.isDesktopApp) {
			new Notice("Power Desk: printing needs the desktop app.");
			return;
		}
		const s = this.plugin.settings;
		const when = (ev: PCEvent) => (ev.allDay ? "All day" : fmtTimeOfMs(ev.startMs, s.use24h, true));

		// Each style reads its own range rather than sharing one: the agenda
		// prints the days you are looking at, and the grid prints the whole
		// month around them, which is not the same set of events even when
		// the grid is what is on screen.
		const agenda: PrintStyle = {
			id: "agenda",
			label: "Agenda",
			hint: "What is on, day by day",
			icon: "list",
			basePt: 10.5,
			landscape: false,
			build: (o) => {
				const win = viewWindow(this.mode, this.anchorKey, s.weekStartsMonday, s.agendaDays, s.dayViewDays);
				const events = this.applyFilter(this.plugin.eventsForWindow(win.fromKey, win.toKey));
				const days: PrintDay[] = [];
				for (let key = win.fromKey; dayDiff(key, win.toKey) >= 0; key = addDays(key, 1)) {
					days.push({
						heading: fmtDayHeading(key),
						events: eventsOnDay(events, key).map((ev) => ({ when: when(ev), title: ev.title, where: ev.location || undefined })),
					});
				}
				return printableAgendaHtml(this.titleEl?.getText() || "Calendar", days, o);
			},
		};

		const month: PrintStyle = {
			id: "month",
			label: "Month grid",
			hint: "The whole month at once",
			icon: "table",
			basePt: 8.5,
			landscape: true,
			build: (o) => {
				const year = +this.anchorKey.slice(0, 4);
				const mon = +this.anchorKey.slice(5, 7) - 1;
				const cells = monthGrid(year, mon, s.weekStartsMonday);
				const events = this.applyFilter(this.plugin.eventsForWindow(cells[0].key, cells[cells.length - 1].key));
				const weeks: PrintCell[][] = [];
				for (let i = 0; i < cells.length; i += 7) {
					weeks.push(
						cells.slice(i, i + 7).map((cell) => ({
							label: String(cell.day),
							dim: !cell.inMonth,
							events: eventsOnDay(events, cell.key).map((ev) => ({ when: when(ev), title: ev.title })),
						}))
					);
				}
				const names = s.weekStartsMonday ? ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"] : ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
				return printableMonthHtml(this.titleEl?.getText() || fmtDayShort(this.anchorKey, true), names, weeks, o);
			},
		};

		new PrintModal(this.app, [agenda, month], style ?? (this.mode === "month" ? "month" : "agenda")).open();
	}

	/** As in the mail view: the other tab's copy of this button has to hear
	 *  about a toggle it did not make. */
	private syncFoldToggle() {
		this.foldToggleBtn?.toggleClass("is-active", this.plugin.focusOn());
	}

	openEventSearch() {
		new EventSearchModal(this.app, this.plugin, (key) => this.goDay(key)).open();
	}

	/** The same window, opened already asking about one person. */
	openEventSearchFor(person: string) {
		new EventSearchModal(this.app, this.plugin, (key) => this.goDay(key), person).open();
	}

	openToolbarEditor() {
		new ToolbarModal(
			this.app,
			this.plugin,
			this.calActions(),
			{
				get: () => this.plugin.settings.calendarToolbar,
				set: (v) => (this.plugin.settings.calendarToolbar = v),
				fallback: DEFAULT_SETTINGS.calendarToolbar,
				leads: "New event",
			},
			() => this.renderCalToolbar()
		).open();
	}

	private renderCalToolbar() {
		const host = this.calToolsEl;
		if (!host) return;
		host.empty();
		const catalog = new Map(this.calActions().map((a) => [a.id, a]));
		for (const id of this.plugin.settings.calendarToolbar) {
			const a = catalog.get(id);
			if (!a) continue;
			const b = host.createEl("button", { cls: "pcal-icon-btn pcal-mail-tool", attr: { "aria-label": a.label } });
			setIcon(b.createSpan(), a.icon);
			b.createSpan({ cls: "pcal-mail-tool-label", text: a.label });
			b.addEventListener("click", (e) => a.run(e));
			// the filter button keeps its own class, since the view marks it
			// active while a filter is on
			if (id === "filter") {
				b.addClass("pcal-filter-btn");
				this.filterBtn = b;
			}
			if (id === "refresh") this.refreshBtn = b;
		}
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
		window.requestAnimationFrame(() => {
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
		window.requestAnimationFrame(() => {
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
		// the card keeps its position once it has one, so this is the way
		// back to landing beside whatever was clicked
		moreItems.push({
			title: "Follow the event again",
			onClick: () => {
				const store = this.app as unknown as { loadLocalStorage: (k: string) => unknown; saveLocalStorage: (k: string, v: unknown) => void };
				const saved = (store.loadLocalStorage("powerdesk:event-card") as { w?: number; h?: number } | null) ?? {};
				store.saveLocalStorage("powerdesk:event-card", { w: saved.w, h: saved.h });
				this.closeCard();
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

		// Sized and placed like the other windows, keeping both. The very
		// first card lands beside the event that was clicked, clamped to the
		// viewport and flipping above when there is no room below; after
		// that it opens where it was left, and "Follow the event again"
		// returns it to landing beside whatever was clicked.
		const stopDrag = makeDraggable(this.app, card, head, "powerdesk:event-card", { w: 340, h: Math.min(420, window.innerHeight - 40) }, () => {
			const r = anchor.getBoundingClientRect();
			const below = r.bottom + 6;
			return {
				x: Math.max(8, Math.min(r.left, window.innerWidth - card.offsetWidth - 8)),
				y: below + card.offsetHeight > window.innerHeight - 8 ? Math.max(8, r.top - card.offsetHeight - 6) : below,
			};
		});

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
			stopDrag();
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
	private descEl!: HTMLElement;
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
		this.modalEl.addClass("pcal-event-window");
		makeMovable(this.app, this, "powerdesk:event-window", { w: 720, h: 700 });
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
			// a real editor, since an agenda is a list and a joining link is a
			// link. Only on a new event: editing one never writes the
			// description back, because only a preview of it was ever fetched
			// and saving that would truncate what is actually there.
			c.createDiv({ cls: "setting-item-name pcal-event-desc-label", text: "Description" });
			const descBar = c.createDiv("pcal-compose-bar");
			this.descEl = c.createDiv({ cls: "pcal-compose-editor pcal-event-desc", attr: { contenteditable: "true" } });
			richToolbar(this.app, descBar, () => this.descEl);
			// an untouched contenteditable still holds a <br>, and sending that
			// as a description gives every invitee an event whose body is one
			// stray tag; empty means empty
			this.descEl.addEventListener("input", () => {
				const hasWords = !!this.descEl.textContent?.trim();
				const hasThings = !!this.descEl.querySelector("img, a, li, hr");
				this.description = hasWords || hasThings ? this.descEl.innerHTML : "";
			});
		}
		const btns = c.createDiv({ cls: "pcal-modal-btns pcal-compose-btns" });
		btns.createEl("button", { text: this.ev ? "Save" : "Create", cls: "mod-cta pcal-send-btn" }).addEventListener("click", () => void this.save());
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Cancel" }).addEventListener("click", () => this.close());
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
				track.setText("No visibility");
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
		c.createDiv({ cls: "pcal-devicecode", text: this.dc.user_code });
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
	/** Held so closing the window takes their floating lists with it: the
	 *  lists live on document.body, outside anything the modal empties. */
	private suggests: AddressSuggest[] = [];
	/** Files to go out with this message, read into memory when picked.
	 *  They upload at send time rather than on pick, so attaching stays
	 *  instant and taking the undo costs nothing. */
	private files: OutgoingFile[] = [];
	private filesEl!: HTMLElement;
	private receipts: Receipts = {};
	private bccInput!: HTMLInputElement;
	/** Which mailbox this goes out from, which the From picker sets. */
	private fromId = "";

	/** Swap the signature when the sending account changes, replacing only
	 *  the block we put there rather than anything typed since. */
	private swapSignature() {
		const old = this.editorEl?.querySelector(".pcal-sig-block");
		const html = this.plugin.signatureHtml(this.fromId, this.opts.mode === "new" ? "new" : "reply");
		if (!old) return;
		const holder = createDiv();
		holder.appendChild(sanitizeHTMLToDom(html));
		old.empty();
		while (holder.firstChild) old.appendChild(holder.firstChild);
	}

	constructor(
		app: App,
		private plugin: PowerDeskPlugin,
		private opts: {
			mode: "new" | "reply" | "replyAll" | "forward";
			mail?: PCMail;
			to?: string;
			subject?: string;
			preferAccountId?: string;
			/** Set when undo put this window back: the fields as they were,
			 *  and the draft that was never sent, so nothing is retyped and
			 *  no second draft is created in the mailbox. */
			resume?: { draft: DraftMessage | null; to: string; cc: string; bcc?: string; subject: string; html: string; files?: OutgoingFile[]; receipts?: Receipts };
		}
	) {
		super(app);
	}

	onOpen() {
		const { mode, mail } = this.opts;
		this.modalEl.addClass("pcal-compose-window");
		this.titleEl.setText(mode === "new" ? "New mail" : mode === "reply" ? "Reply" : mode === "replyAll" ? "Reply all" : "Forward");
		// one remembered geometry for every compose window, since a reply and
		// a new message want the same room and nobody wants to size two
		makeMovable(this.app, this, "powerdesk:compose-window", { w: 900, h: 720 });
		const c = this.contentEl;
		c.addClass("pcal-compose");
		// From: a real picker when several mailboxes can send and this is a
		// fresh message. A reply has no picker on purpose, since its draft
		// already lives in one mailbox and moving it is not a dropdown.
		const senders = this.plugin.settings.graphAccounts.filter((a) => !!a.refresh && a.grantedScope.includes("Mail.Send"));
		const fromRow = c.createDiv("pcal-compose-row pcal-compose-fromrow");
		fromRow.createSpan({ cls: "pcal-compose-label", text: "From" });
		if (mode === "new" && senders.length > 1) {
			const sel = fromRow.createEl("select", { cls: "dropdown" });
			for (const a of senders) sel.createEl("option", { value: a.id, text: this.plugin.nameOf(a) });
			sel.value = this.opts.preferAccountId ?? senders[0].id;
			this.fromId = sel.value;
			sel.addEventListener("change", () => {
				this.fromId = sel.value;
				// the signature belongs to the account, so it changes with it
				this.swapSignature();
			});
		} else {
			const label = mode === "new" ? this.plugin.mailSender(this.opts.preferAccountId)?.label : mail?.accountLabel;
			fromRow.createSpan({ cls: "pcal-compose-fromname", text: label ?? "this mailbox" });
			this.fromId = this.opts.preferAccountId ?? mail?.accountId ?? senders[0]?.id ?? "";
		}

		const row = (label: string, value: string, placeholder?: string): { input: HTMLInputElement; row: HTMLElement } => {
			const r = c.createDiv("pcal-compose-row");
			r.createSpan({ cls: "pcal-compose-label", text: label });
			const inp = r.createEl("input", { attr: { type: "text", spellcheck: "false" } });
			inp.value = value;
			if (placeholder) inp.placeholder = placeholder;
			return { input: inp, row: r };
		};
		const toRow = row("To", this.opts.resume?.to ?? this.opts.to ?? "", "Start typing a name or address");
		this.toInput = toRow.input;
		// Bcc is hidden until wanted, the way every mail client hides it, but
		// it is one click away rather than absent
		const bccToggle = toRow.row.createEl("button", { cls: "pcal-bcc-toggle", text: "Bcc" });
		this.ccInput = row("Cc", this.opts.resume?.cc ?? "").input;
		const bccRow = row("Bcc", this.opts.resume?.bcc ?? "");
		this.bccInput = bccRow.input;
		const showBcc = !!(this.opts.resume?.bcc ?? "").trim();
		bccRow.row.toggle(showBcc);
		bccToggle.toggleClass("is-active", showBcc);
		bccToggle.addEventListener("click", () => {
			const on = !bccRow.row.isShown();
			bccRow.row.toggle(on);
			bccToggle.toggleClass("is-active", on);
			if (on) this.bccInput.focus();
		});
		this.subjInput = row("Subject", this.opts.resume?.subject ?? this.opts.subject ?? "").input;
		this.suggests = [new AddressSuggest(this.toInput, this.plugin), new AddressSuggest(this.ccInput, this.plugin), new AddressSuggest(this.bccInput, this.plugin)];

		const bar = c.createDiv("pcal-compose-bar");
		this.editorEl = c.createDiv({ cls: "pcal-compose-editor", attr: { contenteditable: "true" } });
		richToolbar(this.app, bar, () => this.editorEl, [
			{
				label: "Insert an image",
				icon: "image",
				run: () =>
					new VaultFilePickModal(this.app, (f) => {
						void (async () => {
							const bytes = await this.app.vault.readBinary(f);
							const mime = mimeForExtension(f.extension);
							if (!mime.startsWith("image/")) {
								new Notice("Power Desk: pick an image file.");
								return;
							}
							const src = `data:${mime};base64,${arrayBufferToBase64(bytes)}`;
							insertAtCaret(this.editorEl, createEl("img", { attr: { src, alt: f.basename } }));
						})();
					}).open(),
			},
		]);
		bar.createSpan("pcal-compose-bar-sep");
		richToolbarFull(this.app, bar, () => this.editorEl);
		const res = this.opts.resume;
		if (res) {
			// undo put this back: the body is exactly what was about to go,
			// signature and quoted original already in it
			this.editorEl.appendChild(sanitizeHTMLToDom(res.html));
		} else if (mode === "new") {
			this.editorEl.createEl("p").createEl("br");
			// the signature sits in a block of its own so changing the From
			// account can replace it without touching anything typed
			const sig = this.editorEl.createDiv("pcal-sig-block");
			sig.appendChild(sanitizeHTMLToDom(this.plugin.signatureHtml(this.fromId, "new")));
		} else this.editorEl.createEl("p", { cls: "pcal-compose-loading", text: "Opening the draft..." });

		this.filesEl = c.createDiv("pcal-compose-files");
		if (res?.files?.length) this.files = [...res.files];
		this.renderFiles();

		this.receipts = res?.receipts ?? { read: this.plugin.settings.mailAskReadReceipt, delivery: this.plugin.settings.mailAskDeliveryReceipt };

		const btns = c.createDiv("pcal-modal-btns pcal-compose-btns");
		// Send leads, the way it does in every mail client: it is the thing
		// the window exists for and should not be one grey button among five
		btns.createEl("button", { text: "Send", cls: "mod-cta pcal-send-btn" }).addEventListener("click", () => void this.send());
		btns.createEl("button", { text: "Attach" }).addEventListener("click", (e) => this.attachMenu(e));
		const optBtn = btns.createEl("button", { text: "Options" });
		optBtn.addEventListener("click", (e) => {
			const menu = new Menu();
			const toggle = (title: string, get: () => boolean, set: (v: boolean) => void) =>
				menu.addItem((i) =>
					i
						.setTitle(title)
						.setChecked(get())
						.onClick(() => {
							set(!get());
							optBtn.toggleClass("is-active", !!this.receipts.read || !!this.receipts.delivery);
						})
				);
			toggle("Request a read receipt", () => !!this.receipts.read, (v) => (this.receipts.read = v));
			toggle("Request a delivery receipt", () => !!this.receipts.delivery, (v) => (this.receipts.delivery = v));
			menu.addSeparator();
			menu.addItem((i) => i.setTitle("The recipient's mail app decides whether to answer").setDisabled(true));
			menu.showAtMouseEvent(e);
		});
		optBtn.toggleClass("is-active", !!this.receipts.read || !!this.receipts.delivery);
		// Send later sits beside Send rather than inside a menu on it: it is
		// a different decision, not a variant of the same one
		btns.createEl("button", { text: "Send later" }).addEventListener("click", () => {
			new WhenModal(
				this.app,
				"Send later",
				"The message goes to your mailbox now and the server holds it until this time, so it goes out whether or not Obsidian is open. It waits in Drafts until then, where Outlook can still cancel or edit it.",
				this.plugin.settings.use24h,
				(ms) => void this.send(ms)
			).open();
		});
		btns.createSpan("pcal-compose-btns-gap");
		btns.createEl("button", { text: "Discard" }).addEventListener("click", () => this.close());

		if (res) {
			// the draft the send never used is still in the mailbox; reuse it
			// rather than creating a second one beside it
			this.draft = res.draft;
			window.setTimeout(() => this.editorEl.focus(), 20);
		} else if (mode !== "new" && mail) void this.loadDraft(mail, mode);
		else window.setTimeout(() => this.editorEl.focus(), 20);
	}

	/** Where a file can come from. The vault first: in Obsidian the thing you
	 *  want to send is usually already a note or an asset in it, and picking
	 *  from the vault works identically on every platform, while a file
	 *  dialog does not exist on mobile. */
	private attachMenu(e: MouseEvent) {
		const menu = new Menu();
		menu.addItem((i) =>
			i
				.setTitle("From the vault...")
				.setIcon("folder-open")
				.onClick(() => new VaultFilePickModal(this.app, (f) => void this.addVaultFile(f)).open())
		);
		menu.addItem((i) =>
			i
				.setTitle("From this computer...")
				.setIcon("hard-drive")
				.onClick(() => this.pickDiskFiles())
		);
		menu.showAtMouseEvent(e);
	}

	private async addVaultFile(f: TFile) {
		try {
			const bytes = await this.app.vault.readBinary(f);
			this.files.push({ name: f.name, contentType: mimeForExtension(f.extension), bytes });
			this.renderFiles();
		} catch (err) {
			new Notice("Power Desk: could not read that file. " + (err instanceof Error ? err.message : String(err)));
		}
	}

	/** A file input rather than an Electron dialog, so this works the same on
	 *  the desktop app and on a phone. */
	private pickDiskFiles() {
		const input = createEl("input", { attr: { type: "file", multiple: "true" } });
		input.addClass("pcal-offscreen-input");
		document.body.appendChild(input);
		input.addEventListener("change", () => {
			void (async () => {
				for (const f of Array.from(input.files ?? [])) {
					try {
						this.files.push({ name: f.name, contentType: f.type || mimeForExtension(f.name.split(".").pop() ?? ""), bytes: await f.arrayBuffer() });
					} catch (err) {
						new Notice(`Power Desk: could not read ${f.name}. ` + (err instanceof Error ? err.message : String(err)));
					}
				}
				this.renderFiles();
				input.remove();
			})();
		});
		input.click();
	}

	/** The attached files as chips, with the running total. Outlook refuses
	 *  a message past about 35 MB, so the total is worth showing before the
	 *  send is the thing that tells you. */
	private renderFiles() {
		const host = this.filesEl;
		host.empty();
		if (!this.files.length) return;
		let total = 0;
		this.files.forEach((f, idx) => {
			total += f.bytes.byteLength;
			const chip = host.createDiv("pcal-mail-att pcal-compose-file");
			const badge = attachmentBadge(f.name, f.contentType);
			const ic = chip.createSpan({ cls: "pcal-mail-att-icon", text: badge.label });
			ic.style.setProperty("--pcal-att-color", badge.color);
			const tx = chip.createDiv("pcal-mail-att-text");
			tx.createDiv({ cls: "pcal-mail-att-name", text: f.name });
			tx.createDiv({ cls: "pcal-mail-att-size", text: fmtAttachmentSize(f.bytes.byteLength) });
			const x = chip.createSpan({ cls: "pcal-compose-file-x", attr: { "aria-label": `Remove ${f.name}` } });
			setIcon(x, "x");
			x.addEventListener("click", () => {
				this.files.splice(idx, 1);
				this.renderFiles();
			});
		});
		const note = host.createDiv("pcal-compose-files-total");
		note.setText(`${this.files.length} file${this.files.length === 1 ? "" : "s"}, ${fmtAttachmentSize(total)}`);
		if (total > 30 * 1024 * 1024) {
			note.addClass("is-over");
			note.setText(`${this.files.length} files, ${fmtAttachmentSize(total)}. Most mailboxes refuse a message past about 35 MB.`);
		}
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
		// a reply gets the reply signature, above the quoted original
		this.editorEl.appendChild(sanitizeHTMLToDom(`${this.plugin.signatureHtml(mail.accountId, "reply")}${draft.bodyHtml}`));
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

	/** `whenMs` schedules instead of sending now. A scheduled new message has
	 *  to become a draft first, since the send time is a property of a
	 *  message, so it cannot take the plain sendMail path. */
	private async send(whenMs?: number) {
		const split = (s: string) =>
			s
				.split(/[,;]+/)
				.map((x) => x.trim())
				.filter(Boolean);
		const to = split(this.toInput.value);
		const cc = split(this.ccInput.value);
		const bcc = split(this.bccInput.value);
		const subject = this.subjInput.value.trim() || "(no subject)";
		const html = this.editorEl.innerHTML;
		if (!to.length) {
			new Notice("Power Desk: enter at least one recipient.");
			return;
		}
		// Putting the window back exactly as it was is the whole of undo.
		// The values are taken now rather than read back later: by the time
		// undo runs, this modal has been closed and its DOM emptied.
		const state = { draft: this.draft, to: this.toInput.value, cc: this.ccInput.value, bcc: this.bccInput.value, subject: this.subjInput.value, html, files: [...this.files], receipts: { ...this.receipts } };
		const files = [...this.files];
		const receipts = { ...this.receipts };
		const opts = this.opts;
		const app = this.app;
		const plugin = this.plugin;
		const reopen = () => new RichComposeModal(app, plugin, { ...opts, resume: state }).open();

		const patch = { subject, html, to, cc, bcc, receipts };
		if (this.opts.mode === "new") {
			// plain new mail with nothing hung on it posts straight to
			// sendMail: one request rather than three, and it is the only
			// path the Power Assistant transport can take. An embedded image
			// counts as something hung on it, since it has to become a real
			// attachment and that needs a draft to hang off.
			const plain = !files.length && !whenMs && !/src\s*=\s*["']data:/i.test(html);
			const sender = plain ? this.plugin.mailSender(this.opts.preferAccountId) : null;
			if (plain && !sender) {
				new Notice("Power Desk: no account can send mail.");
				return;
			}
			const fire = async () => {
				if (sender) {
					try {
						await sender.send({ to, cc, bcc, subject, html, receipts });
						new Notice("Power Desk: mail sent.");
					} catch (e) {
						new Notice("Power Desk: " + (e instanceof Error ? e.message : String(e)));
					}
					return;
				}
				await this.plugin.sendComposed({ accountId: null, draftId: null, preferAccountId: this.opts.preferAccountId, patch, files, whenMs });
			};
			// a scheduled send is already its own undo: it waits in Drafts
			// where Outlook can still cancel it, so it goes straight out
			if (whenMs) {
				await fire();
				this.sent = true;
				this.close();
				return;
			}
			this.sent = true;
			this.close();
			this.plugin.holdSend({ label: subject, reopen, fire });
			return;
		}
		if (!this.draft || !this.opts.mail) return;
		const accountId = this.opts.mail.accountId;
		const draftId = this.draft.id;
		if (whenMs) {
			const ok = await this.plugin.sendComposed({ accountId, draftId, patch, files, whenMs });
			if (ok) {
				this.sent = true;
				this.close();
			}
			return;
		}
		this.sent = true;
		this.close();
		this.plugin.holdSend({
			label: subject,
			reopen,
			fire: () => this.plugin.sendComposed({ accountId, draftId, patch, files }).then(() => undefined),
		});
	}

	onClose() {
		if (this.draft && !this.sent && this.opts.mail) void this.plugin.discardMailDraft(this.opts.mail.accountId, this.draft.id);
		for (const s of this.suggests) s.hide();
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
			line("Mail", this.plugin.canMailAccount(a) ? "inbox available in the mail view" : "not granted");
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
			text: "Organizations that block user consent need an admin to grant the delegated Calendars.ReadWrite, Mail.ReadWrite, Mail.Send, MailboxSettings.ReadWrite, ProfilePhoto.Read.All, Contacts.Read, and Tasks.ReadWrite permissions under API permissions; everyone else skips that page entirely.",
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

/** One row of the settings tab. `build` is handed a Setting whose name and
 *  description are already set, so it only adds the controls and any richer
 *  description content. Rows are data rather than drawing code so the two
 *  renderers below cannot disagree about what the tab holds. */
type Row = { name: string; desc?: string; help?: string; cls?: string; aliases?: string[]; build?: (st: Setting) => void | (() => void) };

/** A run of rows under one heading. Each becomes a headed group on 1.13 and
 *  one section div in the fallback. */
type Group = { heading?: string; rows: Row[] };

/** One tab: a native settings page on Obsidian 1.13 and up, a tab button in
 *  the fallback renderer for older builds. */
type Page = { id: string; label: string; groups: Group[] };

/** The controls that earn a row the full width of the two-column settings
 *  layout: anything you type into, drag, or pick from wants the room, while a
 *  toggle or a button is happy beside its neighbour. */
const WIDE_CONTROLS = 'input[type="text"], input[type="password"], input[type="search"], input[type="number"], input[type="range"], textarea, select, .slider';

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
	/** The containers the five source lists draw into. Each list is one row
	 *  owning a container of its own rather than a row per account, for two
	 *  reasons: expanding an account redraws just that list instead of the whole
	 *  tab, which keeps your place on the page; and a list built into the
	 *  definitions would go stale, because reopening a tab renders the cached
	 *  definitions without asking for them again, so an account that arrived
	 *  while settings were closed would be missing. */
	private graphHost: HTMLElement | null = null;
	private googleHost: HTMLElement | null = null;
	private caldavHost: HTMLElement | null = null;
	private icsHost: HTMLElement | null = null;
	private vaultHost: HTMLElement | null = null;

	constructor(
		app: App,
		private plugin: PowerDeskPlugin
	) {
		super(app, plugin);
		// Armed once, for the life of the tab. It used to be set in display() and
		// cleared in hide(), which the declarative renderer would leave null after
		// the first close, since it never calls display() again. refresh() bails
		// when the tab is off screen, so a closed tab still costs nothing.
		plugin.refreshSettingsTab = () => this.refresh();
	}

	hide() {
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

	/** Redraw when the rows themselves change: an account signed in, the shared
	 *  app set, the read-delay slider appearing. Obsidian 1.13 rebuilds the tab
	 *  from getSettingDefinitions(); older builds have only the fallback
	 *  renderer.
	 *
	 *  The plugin calls this whenever an account changes underneath us, so it
	 *  bails when the tab is off screen rather than rebuilding a hidden
	 *  container. */
	refresh() {
		if (!this.containerEl.isShown()) return;
		this.closeHelp(); // whatever the popover is anchored to is about to go
		// update() arrived with the declarative API in 1.13 and minAppVersion is
		// still 1.8.7, so it is reached through a cast rather than named outright:
		// an older build has no definitions to rebuild from and redraws instead.
		const tab = this as unknown as { update?: () => void };
		if (tab.update) tab.update();
		else this.renderFallback();
	}

	/** A small help icon after the setting name carrying the deeper "what does
	 *  this actually do" explanation; hover shows it, a click pins it open. No
	 *  aria-label on the icon or Obsidian's native black tooltip doubles up. */
	private addHelp(st: Setting, text: string) {
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
	}

	/** Obsidian 1.13 and up builds the tab from these and never calls display():
	 *  one native page per tab, standing in for the tab bar the fallback draws
	 *  for older builds. A tab holding more than one section becomes a page of
	 *  headed groups, which is what the headings were doing by hand.
	 *
	 *  Every row renders itself rather than declaring a `control`. A declarative
	 *  control writes through Obsidian's generic setControlValue, which would
	 *  bypass queueSave and the settings merge behind it. */
	getSettingDefinitions(): SettingDefinitionItem[] {
		const pages = this.buildPages();
		const rowsOf = new Map(pages.map((p) => [p.label, p.groups.flatMap((g) => g.rows)] as const));
		return [
			{
				name: "",
				searchable: false, // it is a masthead, not a setting
				render: (st) => {
					st.settingEl.empty();
					this.renderAbout(st.settingEl);
				},
			},
			{
				type: "group",
				search: {
					placeholder: "Search settings...",
					// the entries here are whole tabs, so a tab stays up when anything
					// inside it matches. Obsidian's own search box, top left, reaches
					// the individual settings.
					match: (def, query) => {
						const q = query.trim().toLowerCase();
						if (!q) return true;
						const has = (v: string | undefined) => (v ?? "").toLowerCase().includes(q);
						return (rowsOf.get(def.name) ?? []).some(
							(r) => has(r.name) || has(r.desc) || (r.aliases ?? []).some(has)
						);
					},
				},
				items: pages.map(
					(p): SettingDefinitionPage => ({
						type: "page",
						name: p.label,
						// a tab with one section needs no heading inside it: the page is
						// already named after it. The heading is kept on the group anyway,
						// because the fallback draws every section headed.
						items:
							p.groups.length === 1
								? p.groups[0].rows.map((r) => this.toDefinition(r, p.label))
								: p.groups.map((g) => ({
										type: "group" as const,
										heading: g.heading,
										items: g.rows.map((r) => this.toDefinition(r, p.label)),
									})),
					})
				),
			},
		];
	}

	/** One row as a definition Obsidian can draw. The name and description are
	 *  its to render and it rebuilds both on a redraw, so a row only hands back
	 *  what it hung on the row element itself. */
	private toDefinition(r: Row, page: string): SettingDefinitionRender {
		return {
			name: r.name,
			desc: r.desc,
			// searching the tab name still finds its rows, the way a heading match
			// opened the whole section in the tab bar
			aliases: [...(r.aliases ?? []), page],
			render: (st) => {
				if (r.cls) st.settingEl.addClass(r.cls);
				const teardown = r.build?.(st);
				if (r.help) this.addHelp(st, r.help);
				return teardown;
			},
		};
	}

	/** What this plugin is and which build is running, above the tabs. Read off
	 *  the manifest so it cannot drift from the released version. */
	private renderAbout(el: HTMLElement) {
		el.addClass("pcal-about");
		const head = el.createDiv({ cls: "pcal-about-head" });
		head.createSpan({ cls: "pcal-about-name", text: this.plugin.manifest.name });
		head.createSpan({ cls: "pcal-about-version", text: "v" + this.plugin.manifest.version });
		el.createDiv({ cls: "pcal-about-desc", text: this.plugin.manifest.description });
		// One Buy Me a Coffee page serves every Power Plugin, and a payment says
		// nothing about which one it came from, nor about what the person wanted.
		// The note that rides along can carry both, so it asks for both. The name is
		// read from the manifest rather than written out here, so it cannot drift
		// from what the plugin is actually called.
		//
		// It invites a request without promising to build one. What can be built
		// depends on what the mailbox and vault APIs allow, and a promise broken at
		// the price of a coffee would cost more than never making it. The last
		// sentence points at what has already happened instead, which is true and
		// commits to nothing.
		const support = el.createDiv({ cls: "pcal-about-support" });
		support.createEl("a", { text: "Buy me a coffee", href: "https://buymeacoffee.com/powerplugins" });
		support.createSpan({
			text: `. One page covers every Power Plugin, so mention ${this.plugin.manifest.name} in the note, and say what would make it better while you are there. A good deal of what is in these plugins started as someone's note.`,
		});
	}

	/** The pre-1.13 renderer: every section on one page, with a tab bar and a
	 *  search box of our own because there was no declarative API to hand the
	 *  work to. Obsidian 1.13 and up ignores this and renders the definitions
	 *  above instead, so the two only ever differ in how they draw, never in
	 *  what they draw. */
	display() {
		this.renderFallback();
	}

	private renderFallback() {
		const root = this.containerEl;
		root.empty();
		this.closeHelp(); // a re-render orphans any popover anchored to the old DOM

		const pages = this.buildPages();
		if (!pages.some((p) => p.id === this.activeTab)) this.activeTab = pages[0].id;

		this.renderAbout(root.createDiv({ cls: "pcal-about-standalone" }));

		const searchWrap = root.createDiv({ cls: "pcal-settings-search" });
		const searchInput = searchWrap.createEl("input", { cls: "pcal-settings-search-input" });
		searchInput.type = "search";
		searchInput.placeholder = "Search settings...";
		searchInput.value = this.query;

		const tabBar = root.createDiv({ cls: "pcal-settings-tabs" });
		const body = root.createDiv({ cls: "pcal-settings-body" });

		// one section div per group, tagged with its tab so the tab bar and the
		// search box below can show and hide whole sections at a time
		for (const p of pages) {
			for (const [i, g] of p.groups.entries()) {
				const sec = body.createDiv({ cls: "pcal-settings-section" });
				sec.dataset.tab = p.id;
				sec.dataset.name = (g.heading ?? p.label).toLowerCase();
				if (i === 0) sec.dataset.first = "1"; // the tab's first heading skips its top border
				if (g.heading) new Setting(sec).setName(g.heading).setHeading();
				for (const r of g.rows) this.drawRow(sec, r);
			}
		}

		const setVisible = (el: HTMLElement, v: boolean) => el.toggleClass("pcal-hidden", !v);
		const applyView = () => {
			const q = this.query.trim().toLowerCase();
			setVisible(tabBar, !q);
			for (const sec of Array.from(body.children) as HTMLElement[]) {
				const items = Array.from(sec.querySelectorAll<HTMLElement>(":scope > .setting-item:not(.setting-item-heading)"));
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
					const hit = nameHit || name.includes(q) || desc.includes(q) || (it.dataset.pcalAlias ?? "").includes(q);
					setVisible(it, hit);
					if (hit) anyHit = true;
				}
				setVisible(sec, anyHit);
			}
		};

		for (const p of pages) {
			const btn = tabBar.createEl("button", { text: p.label, cls: "pcal-settings-tab" });
			btn.toggleClass("is-active", p.id === this.activeTab);
			btn.onclick = () => {
				if (this.activeTab === p.id) return;
				this.activeTab = p.id;
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

	/** One row into a container, in the order Obsidian applies a definition:
	 *  name and description first, then the row's own content, so a row that
	 *  appends to either element lands in the same place under both renderers. */
	private drawRow(into: HTMLElement, r: Row) {
		const st = new Setting(into).setName(r.name);
		if (r.desc) st.setDesc(r.desc);
		if (r.cls) st.settingEl.addClass(r.cls);
		if (r.aliases?.length) st.settingEl.dataset.pcalAlias = r.aliases.join(" ").toLowerCase();
		r.build?.(st);
		if (r.help) this.addHelp(st, r.help);
		// A row with something to type in, drag, or choose from takes the full
		// width of the two-column layout; a row with a switch or a button sits in
		// one column. The control is built by then, so the row can be asked once
		// here instead of the stylesheet asking it forever with :has().
		if (st.settingEl.querySelector(WIDE_CONTROLS)) st.settingEl.addClass("pcal-wide");
	}

	/** A row that owns a container instead of a control: the source lists draw
	 *  themselves into it and redraw in place, so adding or expanding an account
	 *  never rebuilds the tab. */
	private listRow(name: string, aliases: string[], take: (host: HTMLElement | null) => void, draw: () => void): Row {
		return {
			name: "",
			aliases: [name, ...aliases],
			build: (st) => {
				st.settingEl.empty();
				st.settingEl.addClass("pcal-list-host");
				take(st.settingEl.createDiv({ cls: "pcal-list" }));
				draw();
				return () => take(null);
			},
		};
	}

	/** The Microsoft accounts, each expanding to its name, inbox, and calendars. */
	private drawGraphAccounts() {
		const host = this.graphHost;
		if (!host) return;
		host.empty();
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		for (const a of s.graphAccounts) {
			const missing: string[] = [];
			if (a.refresh && !this.plugin.canWriteAccount(a)) missing.push("editing");
			if (a.refresh && !this.plugin.canMailAccount(a)) missing.push("mail");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("Mail.ReadWrite")) missing.push("reply windows");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("MailboxSettings")) missing.push("inbox rules");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("ProfilePhoto")) missing.push("sender photos");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("Contacts")) missing.push("saved contacts");
			if (a.refresh && this.plugin.canMailAccount(a) && !a.grantedScope.includes("Tasks")) missing.push("tasks");
			const open = this.expandedAccounts.has(a.id);
			const row = new Setting(host)
				.setName(this.plugin.nameOf(a))
				.setDesc((a.nickname?.trim() ? a.label + " · " : "") + "Microsoft 365" + (a.clientId.trim() && a.clientId.trim() !== this.plugin.effectiveClientId() ? " · own app" : "") + (!a.refresh ? " · signed out" : missing.length ? ` · reconnect to enable ${missing.join(" and ")}` : ""))
				.setClass("pcal-account-head");
			row.addExtraButton((b) =>
				b.setIcon(open ? "chevron-down" : "chevron-right").setTooltip(open ? "Hide details" : "Rename, inbox, calendars").onClick(() => {
					if (open) this.expandedAccounts.delete(a.id);
					else this.expandedAccounts.add(a.id);
					this.drawGraphAccounts();
				})
			);
			// any missing power offers the reconnect, not just missing write access
			if (!a.refresh || missing.length) row.addButton((b) => b.setButtonText("Reconnect").setCta().onClick(() => void this.plugin.connectGraph(a)));
			if (a.refresh)
				row.addButton((b) =>
					b.setButtonText("Refresh calendars").onClick(() => {
						void this.plugin.syncGraphCalendars(a).then(() => {
							this.plugin.sourcesChanged();
							this.drawGraphAccounts();
						});
					})
				);
			row.addButton((b) => markDestructive(b.setButtonText("Remove")).onClick(() => this.plugin.removeGraphAccount(a)));
			if (!open) continue;
			const nameSt = new Setting(host)
				.setName("Name")
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
						this.drawGraphAccounts();
					});
				});
			this.addHelp(nameSt, "What this account is called throughout the plugin: in the calendar's source list, on mail, and in the account picker. Handy when two accounts share a provider (work and personal Microsoft) and the addresses are the only thing telling them apart. Leaving it empty falls back to the address itself.");
			if (this.plugin.canMailAccount(a)) {
				const mailSt = new Setting(host)
					.setName("Inbox in the mail view")
					.setClass("pcal-subsetting")
					.addToggle((t) =>
						t.setValue(a.mail !== false).onChange((v) => {
							a.mail = v;
							save();
							this.plugin.mailChanged();
						})
					);
				this.addHelp(mailSt, "Whether this account's inbox appears in the mail view. Turning it off leaves the account's calendars syncing normally and simply stops its mail from showing, which is what you want for an account you read elsewhere.");
			}
			for (const cal of a.calendars) {
				new Setting(host)
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
	}

	/** The Google accounts, each expanding to its name and calendars. */
	private drawGoogleAccounts() {
		const host = this.googleHost;
		if (!host) return;
		host.empty();
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		for (const g of s.googleAccounts) {
			const open = this.expandedAccounts.has(g.id);
			const row = new Setting(host)
				.setName(this.plugin.nameOf(g))
				.setDesc((g.nickname?.trim() ? g.label + " · " : "") + "Google" + (g.refresh ? "" : " · signed out"))
				.setClass("pcal-account-head");
			row.addExtraButton((b) =>
				b.setIcon(open ? "chevron-down" : "chevron-right").setTooltip(open ? "Hide details" : "Rename, calendars").onClick(() => {
					if (open) this.expandedAccounts.delete(g.id);
					else this.expandedAccounts.add(g.id);
					this.drawGoogleAccounts();
				})
			);
			if (!g.refresh) row.addButton((b) => b.setButtonText("Reconnect").setCta().onClick(() => void this.plugin.connectGoogle(g)));
			else
				row.addButton((b) =>
					b.setButtonText("Refresh calendars").onClick(() => {
						void this.plugin.syncGoogleCalendars(g).then(() => {
							this.plugin.sourcesChanged();
							this.drawGoogleAccounts();
						});
					})
				);
			row.addButton((b) => markDestructive(b.setButtonText("Remove")).onClick(() => this.plugin.removeGoogleAccount(g)));
			if (!open) continue;
			const nameSt = new Setting(host)
				.setName("Name")
				.setDesc("A friendly name shown wherever this account appears; empty keeps the address.")
				.setClass("pcal-subsetting")
				.addText((t) => {
					t.setPlaceholder(g.label).setValue(g.nickname ?? "").onChange((v) => {
						g.nickname = v;
						save();
					});
					t.inputEl.addEventListener("blur", () => {
						this.plugin.sourcesChanged();
						this.drawGoogleAccounts();
					});
				});
			this.addHelp(nameSt, "What this Google account is called in the calendar's source list and the account picker. Empty keeps the address.");
			for (const cal of g.calendars) {
				new Setting(host)
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
	}

	/** The CalDAV accounts and the collections under each. */
	private drawCaldavAccounts() {
		const host = this.caldavHost;
		if (!host) return;
		host.empty();
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		for (const account of s.caldavAccounts) {
			new Setting(host)
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
							this.drawCaldavAccounts();
						}).open();
					})
				)
				.addButton((b) =>
					markDestructive(b.setButtonText("Remove")).onClick(() => {
						s.caldavAccounts = s.caldavAccounts.filter((a) => a.id !== account.id);
						save();
						this.plugin.sourcesChanged();
						this.drawCaldavAccounts();
					})
				);
			for (const coll of account.collections) {
				new Setting(host)
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
	}

	/** The ICS subscriptions. */
	private drawIcsFeeds() {
		const host = this.icsHost;
		if (!host) return;
		host.empty();
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		for (const feed of s.icsFeeds) {
			new Setting(host)
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
							this.drawIcsFeeds();
						}).open();
					})
				)
				.addButton((b) =>
					markDestructive(b.setButtonText("Remove")).onClick(() => {
						s.icsFeeds = s.icsFeeds.filter((f) => f.id !== feed.id);
						save();
						this.plugin.sourcesChanged();
						this.drawIcsFeeds();
					})
				);
		}
	}

	/** The vault-note sources. */
	private drawVaultSources() {
		const host = this.vaultHost;
		if (!host) return;
		host.empty();
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		for (const v of s.vaultSources) {
			new Setting(host)
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
							this.drawVaultSources();
						}).open();
					})
				)
				.addButton((b) =>
					markDestructive(b.setButtonText("Remove")).onClick(() => {
						s.vaultSources = s.vaultSources.filter((x) => x.id !== v.id);
						save();
						this.plugin.sourcesChanged();
						this.drawVaultSources();
					})
				);
		}
	}

	/** Every row of the settings tab, in order, as plain data: the one source
	 *  both renderers draw from, so they cannot drift apart. Built fresh on each
	 *  render because much of this tab reflects live account state. */
	private buildPages(): Page[] {
		const s = this.plugin.settings;
		const save = () => this.plugin.queueSave();
		// A section's opening paragraph, carried as a row rather than loose text:
		// the theme cards every .setting-item, so a bare <p> floats outside the
		// boxes and breaks the column the rest of the rows line up on.
		const intro = (text: string): Row => ({ name: "", desc: text, cls: "pcal-section-intro" });

		/* ---------------- Microsoft 365 ---------------- */

		const graphAccounts: Row[] = [
			intro("Work or personal Microsoft accounts, signed in with a device code in your browser; the plugin never sees a password. Each account expands to its calendars, and a friendly name tells work and personal apart at a glance."),
			{
				name: "Add a Microsoft 365 account",
				desc: "Work or personal; a short wizard picks the right app registration and signs it in.",
				help: "Starts the sign-in wizard. You sign in on Microsoft's own page in your browser and paste back a short code, so no password ever reaches the plugin. A work account usually rides a shared app registration; a personal one gets its own. Each account you add brings its calendars along and can be renamed afterwards.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Add account").setCta().onClick(() => new GraphAccountWizard(this.app, this.plugin).open()));
				},
			},
			this.listRow(
				"Microsoft 365 accounts",
				["calendars", "inbox", "rename", "reconnect"],
				(host) => (this.graphHost = host),
				() => this.drawGraphAccounts()
			),
		];

		const graphAppReady = !!this.plugin.effectiveClientId();
		const graphApp: Row[] = [
			intro("The shared app registration new work sign-ins ride, borrowed from Power Assistant when present. Add account handles all of this on its own; this section is for inspecting or swapping the shared app. Already-connected accounts keep the app they signed in with."),
			{
				name: "Create the app registration",
				desc: graphAppReady
					? this.plugin.usingSiblingApp()
						? "Power Assistant's app registration is borrowed automatically; nothing to do here."
						: "A shared app is set; new sign-ins use it."
					: "The Add account wizard creates this inline; these steps set the shared app from here instead.",
				help: "The registration is a one-time entry in Microsoft's systems that sign-ins name, so consent screens can say who is asking. The guide covers the registration fields, the one switch device-code sign-in needs, and where the Application (client) ID lives: paste that ID below when you have it. When Power Assistant is installed and already set up, its app is borrowed automatically and none of this is needed.",
				build: (st) => {
					st.addButton((b) => {
						b.setButtonText("Show the steps").onClick(() =>
							new AzureAppGuideModal(this.app, (id) => {
								s.graphClientId = id;
								save();
								this.refresh();
							}).open()
						);
						if (!graphAppReady) b.setCta();
					});
				},
			},
			{
				name: "Shared app",
				desc: this.plugin.usingSiblingApp()
					? "Power Assistant's app registration, borrowed automatically."
					: s.graphClientId.trim()
						? s.graphClientId.trim() + (s.graphTenant.trim() ? " · tenant " + s.graphTenant.trim() : "")
						: "None yet; the wizard sets it on the first work sign-in.",
				help: "The Azure app registration that new work sign-ins use. Power Assistant's is borrowed automatically when that plugin is present, so there is usually nothing to set. Accounts already connected keep whichever app they signed in with, so swapping this only affects the next sign-in.",
				build: (st) => {
					st.addExtraButton((b) =>
						b.setIcon("pencil").setTooltip(this.showGraphAppFields ? "Hide the fields" : "Edit the ID and tenant").onClick(() => {
							this.showGraphAppFields = !this.showGraphAppFields;
							this.refresh();
						})
					);
				},
			},
		];
		if (this.showGraphAppFields) {
			graphApp.push({
				name: "Application (client) ID",
				cls: "pcal-subsetting",
				desc: this.plugin.usingSiblingApp() ? "Using Power Assistant's app registration. Enter an ID here to use a different one." : "From your Azure app registration's Overview page.",
				help: "From the Overview page of an app created with 'Show the steps' above (any registration with public client flows on works). When Power Assistant is installed and already set up, its app is borrowed automatically and this field can stay empty. Already-connected accounts keep the app they signed in with; this field only steers new sign-ins.",
				build: (st) => {
					st.addText((t) => {
						t.setPlaceholder(this.plugin.usingSiblingApp() ? this.plugin.effectiveClientId() : "00000000-0000-...").setValue(s.graphClientId).onChange((v) => {
							s.graphClientId = v;
							save();
						});
						t.inputEl.addEventListener("blur", () => this.refresh());
					});
				},
			});
			graphApp.push({
				name: "Tenant",
				cls: "pcal-subsetting",
				desc: "'common' works for most accounts; single-organization apps need their Directory (tenant) ID.",
				help: "Which Microsoft directory the sign-in goes through. 'common' suits personal accounts and most work ones; an organization whose app registration is single-tenant needs its Directory (tenant) ID instead, which an administrator can supply. Getting this wrong shows up as a sign-in that refuses the account rather than anything subtler.",
				build: (st) => {
					st.addText((t) => {
						t.setPlaceholder(this.plugin.usingSiblingApp() && this.plugin.effectiveTenant() !== "common" ? this.plugin.effectiveTenant() : "common").setValue(s.graphTenant).onChange((v) => {
							s.graphTenant = v;
							save();
						});
						t.inputEl.addEventListener("blur", () => this.refresh());
					});
				},
			});
		}

		/* ---------------- Google ---------------- */

		const googleAccounts: Row[] = [
			intro("Google accounts, signed in through your browser. Each account expands to its calendars, and a friendly name tells accounts apart at a glance."),
			{
				name: "Add a Google account",
				desc: this.plugin.googleReady() ? "Signs in through your browser. Desktop only; the connection syncs." : "Needs your Google Cloud client ID and secret first (Google app, below).",
				help: "Starts Google's sign-in in your browser. Google requires your own OAuth client (the id and secret above) because installed apps cannot ship a shared one; once that is filled in, adding accounts is a click each.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Add account").setCta().onClick(() => void this.plugin.connectGoogle()));
				},
			},
			this.listRow(
				"Google accounts",
				["calendars", "rename", "reconnect"],
				(host) => (this.googleHost = host),
				() => this.drawGoogleAccounts()
			),
		];

		const googleApp: Row[] = [
			intro("Google requires your own free Google Cloud project, since its terms forbid shipping shared credentials inside an open plugin. One time: create a project at console.cloud.google.com, enable the Google Calendar API, configure the OAuth consent screen as External and press Publish app, then create an OAuth client of type Desktop app and paste its ID and secret here. The README walks through every step."),
			{
				name: "Client ID",
				help: "Press Publish app on the OAuth consent screen: a project left in Testing mode gets refresh tokens that expire every 7 days, which reads as being signed out weekly. Publishing an External app you never submit for verification just means a one-time 'unverified app' warning at sign-in.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder("....apps.googleusercontent.com").setValue(s.googleClientId).onChange((v) => {
							s.googleClientId = v;
							save();
						})
					);
				},
			},
			{
				name: "Client secret",
				help: "The secret half of your Google OAuth client, from the same credentials page as the client id. Google documents the installed-app secret as not confidential, which is why sharing it across your own devices is how they all sign in to the same app. Stored per device and sent only to Google.",
				build: (st) => {
					st.addText((t) => {
						t.inputEl.type = "password";
						t.setPlaceholder("GOCSPX-...").setValue(s.googleClientSecret).onChange((v) => {
							s.googleClientSecret = v;
							save();
						});
					});
				},
			},
		];

		/* ---------------- CalDAV ---------------- */

		const caldav: Row[] = [
			intro("iCloud, Fastmail, Nextcloud, Radicale, and anything else speaking CalDAV. Read-only. Credentials stay in this vault's plugin data; iCloud and Fastmail want an app-specific password, not your account password."),
			this.listRow(
				"CalDAV accounts",
				["collections", "calendars"],
				(host) => (this.caldavHost = host),
				() => this.drawCaldavAccounts()
			),
			{
				name: "Add a CalDAV account",
				help: "Connects a calendar server that speaks CalDAV (Fastmail, iCloud, Nextcloud, Radicale, and most self-hosted servers). You provide the server URL and credentials; an app-specific password is the right choice wherever the provider offers one.",
				build: (st) => {
					st.addButton((b) =>
						b.setButtonText("Add account").onClick(() => {
							new CaldavAccountModal(this.app, null, (account) => {
								s.caldavAccounts = [...s.caldavAccounts, account];
								save();
								this.plugin.sourcesChanged();
								this.drawCaldavAccounts();
							}).open();
						})
					);
				},
			},
		];

		/* ---------------- ICS feeds ---------------- */

		const ics: Row[] = [
			intro("Read-only iCalendar subscriptions: holiday calendars, a published Outlook or Google calendar, team schedules."),
			this.listRow(
				"ICS feeds",
				["subscriptions", "calendars"],
				(host) => (this.icsHost = host),
				() => this.drawIcsFeeds()
			),
			{
				name: "Add an ICS feed",
				help: "Subscribes to a read-only calendar published as an .ics URL: a shared team calendar, a sports schedule, a holiday feed. Feeds are fetched on the refresh timer and never written to, so nothing here can change the source.",
				build: (st) => {
					st.addButton((b) =>
						b.setButtonText("Add feed").onClick(() => {
							new IcsFeedModal(this.app, null, (feed) => {
								s.icsFeeds = [...s.icsFeeds, feed];
								save();
								this.plugin.sourcesChanged();
								this.drawIcsFeeds();
							}).open();
						})
					);
				},
			},
		];

		/* ---------------- Vault notes ---------------- */

		const vault: Row[] = [
			intro("Notes with a date property render as events beside your calendars: deadlines, dated meeting notes, anything. Click one to open the note; drag a timed one to rewrite its date. Entirely local."),
			this.listRow(
				"Vault notes",
				["sources", "date property"],
				(host) => (this.vaultHost = host),
				() => this.drawVaultSources()
			),
			{
				name: "Add a vault source",
				help: "Turns notes in your vault into calendar events, reading a date property you choose. Useful for anything already tracked as notes (birthdays, deadlines, travel) that you would rather see on the calendar than remember separately.",
				build: (st) => {
					st.addButton((b) =>
						b.setButtonText("Add source").onClick(() => {
							new VaultSourceModal(this.app, null, (src) => {
								s.vaultSources = [...s.vaultSources, src];
								save();
								this.plugin.sourcesChanged();
								this.drawVaultSources();
							}).open();
						})
					);
				},
			},
		];

		/* ---------------- Calendar ---------------- */

		const views: Row[] = [
			{
				name: "Default view",
				help: "Which layout the calendar opens on: month for the shape of the weeks ahead, week or work week for hour-by-hour detail, day for one column, agenda for a plain chronological list. Switching views in the calendar itself does not change this, so the view you open on stays predictable.",
				build: (st) => {
					st.addDropdown((d) =>
						d
							.addOptions({ month: "Month", week: "Week", workweek: "Work week", day: "Day", agenda: "Agenda" })
							.setValue(s.defaultMode)
							.onChange((v) => {
								s.defaultMode = v as ViewMode;
								save();
							})
					);
				},
			},
			{
				name: "Default view on phones",
				help: "The view a phone opens on, kept separate because a month grid that reads well on a monitor is cramped on a handset. Agenda is usually the right answer there.",
				build: (st) => {
					st.addDropdown((d) =>
						d
							.addOptions({ agenda: "Agenda", day: "Day", month: "Month", week: "Week", workweek: "Work week" })
							.setValue(s.phoneDefaultMode)
							.onChange((v) => {
								s.phoneDefaultMode = v as ViewMode;
								save();
							})
					);
				},
			},
			{
				name: "Week starts on",
				help: "Which day begins the week in the month grid and week view. Affects only how the calendar is drawn, never the events themselves.",
				build: (st) => {
					st.addDropdown((d) =>
						d
							.addOptions({ monday: "Monday", sunday: "Sunday" })
							.setValue(s.weekStartsMonday ? "monday" : "sunday")
							.onChange((v) => {
								s.weekStartsMonday = v === "monday";
								save();
								this.plugin.notify();
							})
					);
				},
			},
			{
				name: "24-hour clock",
				help: "Show times as 14:00 rather than 2 PM, everywhere the plugin prints a time.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.use24h).onChange((v) => {
							s.use24h = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Week numbers",
				help: "Show the ISO week number beside each week in the month and week views. Weeks belong to the year containing their Thursday, so the first days of January can carry the previous year's final week number.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.showWeekNumbers).onChange((v) => {
							s.showWeekNumbers = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Agenda covers",
				desc: `${s.agendaDays} days`,
				help: "How far ahead the agenda view lists, in days. A longer window means one scroll shows more of what is coming, at the cost of a busier list.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(7, 90, 1)
							.setValue(s.agendaDays)
							.onChange((v) => {
								s.agendaDays = v;
								st.setDesc(`${v} days`);
								save();
								this.plugin.notify();
							})
					);
				},
			},
			{
				name: "Scroll the day to",
				desc: `${s.dayStartHour}:00`,
				help: "Where the week and day grids scroll to when they open. The whole 24 hours stay reachable; this only picks the first hour in view.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(0, 12, 1)
							.setValue(s.dayStartHour)
							.onChange((v) => {
								s.dayStartHour = v;
								st.setDesc(`${v}:00`);
								save();
							})
					);
				},
			},
			{
				name: "Second time zone",
				desc: "Shown beside the hours in week and day views.",
				help: "An IANA zone name like Europe/Berlin, Asia/Manila, or America/Los_Angeles. Leave empty for one clock. An unrecognized name simply hides the column rather than erroring.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder("Europe/Berlin").setValue(s.secondTimeZone).onChange((v) => {
							s.secondTimeZone = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Show declined events",
				help: "Only Microsoft 365 knows which invites you declined. Off, they disappear; on, they render struck through.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.showDeclined).onChange((v) => {
							s.showDeclined = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Awaiting response color",
				help: "Background tint for invites you have not accepted yet.",
				build: (st) => {
					st.addColorPicker((p) =>
						p.setValue(s.calNeedsActionColor).onChange((v) => {
							s.calNeedsActionColor = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
		];

		const refreshRows: Row[] = [
			{
				name: "Meeting reminders",
				desc: s.reminderMinutes > 0 ? `${s.reminderMinutes} minutes before` : "Off",
				help: "While Obsidian is open, a sticky notice appears before each timed meeting, with a Join button when there is a link. Declined meetings and all-day events stay quiet.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(0, 30, 1)
							.setValue(s.reminderMinutes)
							.onChange((v) => {
								s.reminderMinutes = v;
								st.setDesc(v > 0 ? `${v} minutes before` : "Off");
								save();
							})
					);
				},
			},
			{
				name: "Auto-refresh",
				desc: s.refreshMinutes > 0 ? `Every ${s.refreshMinutes} minute${s.refreshMinutes === 1 ? "" : "s"}` : "Manual only (R or the refresh button)",
				help: "How often connected calendars and mail are re-fetched while Obsidian is open. Shorter means fresher and more requests; 0 turns the timer off so nothing is fetched until you refresh by hand.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(0, 60, 1)
							.setValue(s.refreshMinutes)
							.onChange((v) => {
								s.refreshMinutes = v;
								st.setDesc(v > 0 ? `Every ${v} minute${v === 1 ? "" : "s"}` : "Manual only (R or the refresh button)");
								save();
								this.plugin.refreshCadenceChanged();
							})
					);
				},
			},
		];

		const availability: Row[] = [
			intro("Copy free slots (command palette) reads this window and puts the next five workdays' open times on the clipboard, ready to paste into a mail."),
			{
				name: "Free day starts",
				desc: `${s.freeFromHour}:00`,
				help: "The earliest hour the free-slot finder will suggest. It never proposes a meeting before this, so an early calendar block does not turn into a breakfast invitation.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(5, 12, 1)
							.setValue(s.freeFromHour)
							.onChange((v) => {
								s.freeFromHour = v;
								st.setDesc(`${v}:00`);
								save();
							})
					);
				},
			},
			{
				name: "Free day ends",
				desc: `${s.freeToHour}:00`,
				help: "The latest hour the free-slot finder will suggest, so a gap late in the evening is not offered as availability.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(12, 22, 1)
							.setValue(s.freeToHour)
							.onChange((v) => {
								s.freeToHour = v;
								st.setDesc(`${v}:00`);
								save();
							})
					);
				},
			},
		];

		const weather: Row[] = [
			intro("The sidebar agenda's day headers carry the forecast from Open-Meteo, a free service needing no account. Type a city and press Look up; clearing the city and looking up again switches weather off."),
			{
				name: "Location",
				desc: s.weatherLat.trim() ? `Forecast for ${s.weatherPlace || `${s.weatherLat}, ${s.weatherLon}`}.` : "A city name, like 'Chicago, IL' or 'Fort Wayne'.",
				help: "The place the forecast is for. Type a town or city and pick from the suggestions; the coordinates are stored, so a renamed or ambiguous place stays resolved to the spot you chose.",
				build: (st) => {
					let weatherQuery = s.weatherPlace;
					st.addText((t) => t.setValue(s.weatherPlace).onChange((v) => (weatherQuery = v)));
					st.addButton((b) =>
						b.setButtonText("Look up").setCta().onClick(async () => {
							const q = weatherQuery.trim();
							if (!q) {
								s.weatherPlace = s.weatherLat = s.weatherLon = "";
								this.plugin.clearWeather();
								save();
								this.refresh();
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
							this.refresh();
							new Notice(`Power Desk: forecast set to ${hit.label}.`);
						})
					);
				},
			},
			{
				name: "Unit",
				help: "Whether the forecast shows Fahrenheit or Celsius.",
				build: (st) => {
					st.addDropdown((d) =>
						d
							.addOptions({ f: "Fahrenheit", c: "Celsius" })
							.setValue(s.weatherUnit)
							.onChange((v) => {
								s.weatherUnit = v as "f" | "c";
								this.plugin.clearWeather();
								save();
							})
					);
				},
			},
		];

		/* ---------------- Mail ---------------- */

		const mail: Row[] = [
			{
				name: "Mail history",
				desc: `Pull the last ${s.mailHistoryDays} days of mail. Also sets how far back Power Assistant's "Ask your email" can reach, since it searches only what is cached here.`,
				help: "How far back mail is pulled. This is also the ceiling on what Power Assistant's 'Ask your email' can search, because that window only ever indexes messages this plugin has already fetched. Raising it makes the next sync fetch more, once.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(7, 365, 1)
							.setValue(Math.min(365, Math.max(7, s.mailHistoryDays || 45)))
							.onChange((v) => {
								s.mailHistoryDays = v;
								st.setDesc(`Pull the last ${v} days of mail. Also sets how far back Power Assistant's "Ask your email" can reach, since it searches only what is cached here.`);
								save();
							})
					);
				},
			},
			{
				name: "Messages kept per folder",
				desc: `Retain up to ${s.mailMaxMessages} of the newest messages in each folder. Higher means deeper search and more memory; the initial sync fetches more the first time.`,
				help: "How many of the newest messages survive each sync, per folder. This is the setting that actually bounds how much mail you can search: a wide day range changes nothing while this stays low, because older messages are dropped no matter how far back the window reaches. Higher costs memory and a longer first sync.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(50, 5000, 50)
							.setValue(Math.min(5000, Math.max(50, s.mailMaxMessages || 50)))
							.onChange((v) => {
								s.mailMaxMessages = v;
								st.setDesc(`Retain up to ${v} of the newest messages in each folder. Higher means deeper search and more memory; the initial sync fetches more the first time.`);
								save();
							})
					);
				},
			},
			{
				name: "Saved-mail folder",
				desc: `Where the "Save to note" button files an email. Empty uses the calendar notes folder (${s.notesFolder.trim() || "Calendar"}). Point it at a Power Connect encrypted folder to keep saved mail encrypted on Dropbox.`,
				help: "Where the Save-to-note button files an email. It has its own setting rather than sharing the calendar's notes folder so that filing mail somewhere specific (a Power Connect encrypted folder, say) does not drag event notes along with it. Empty falls back to the calendar notes folder.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder(s.notesFolder.trim() || "Calendar").setValue(s.mailNotesFolder).onChange((v) => {
							s.mailNotesFolder = v.trim();
							save();
						})
					);
				},
			},
			{
				name: "Keep the pictures",
				cls: "pcal-subsetting",
				desc: "Bring a saved email's inline images into the vault, into an attachments folder beside the note.",
				help: "A saved message is converted to Markdown, so its headings, lists and links come with it. Its pictures cannot: they live in the mailbox as attachments the message points at, and a note that points at those shows nothing. On, each one is written into an attachments folder under the saved-mail folder and embedded in the note, which keeps it inside a Power Connect protected folder rather than in the vault's general attachment folder, and keeps the note and its pictures together when either moves. Tracking pixels are never saved, and images the sender hosts elsewhere stay links, exactly as the message has them. Off, the note keeps the words and nothing is written but the note.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailNoteImages).onChange((v) => {
							s.mailNoteImages = v;
							save();
						})
					);
				},
			},
			{
				name: "Mark as read",
				help: "When a message counts as read. 'As soon as it is selected' matches most mail apps; a short delay avoids marking things read as you arrow past them; 'only when I mark it myself' leaves the decision entirely to you.",
				build: (st) => {
					st.addDropdown((d) =>
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
								this.refresh(); // the seconds slider follows the choice
							})
					);
				},
			},
		];
		if (s.markRead === "delay") {
			mail.push({
				name: "Mark read after",
				cls: "pcal-subsetting",
				desc: `${s.markReadSeconds} second${s.markReadSeconds === 1 ? "" : "s"}`,
				help: "How long a message stays selected before it counts as read. Long enough to arrow through a list without clearing everything, short enough that a message you actually stopped on is marked.",
				build: (st) => {
					st.addSlider((sl) =>
						showSliderValue(sl)
							.setLimits(1, 30, 1)
							.setValue(s.markReadSeconds)
							.onChange((v) => {
								s.markReadSeconds = v;
								st.setDesc(`${v} second${v === 1 ? "" : "s"}`);
								save();
							})
					);
				},
			});
		}
		mail.push(
			{
				name: "Unread filter keeps items unread",
				help: "While the Unread filter is on in the mail view, selecting a message never marks it read; only the explicit read buttons do. Exactly Outlook's 'always keep items unread' behavior for unread filtering.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.unreadFilterKeepsUnread).onChange((v) => {
							s.unreadFilterKeepsUnread = v;
							save();
						})
					);
				},
			},
			{
				name: "Split the inbox",
				desc: "Group the inbox into Priority, Focused, Notifications, and Other, the way Spark and Shortwave bundle a list. Folders you file into stay flat.",
				help: "Priority is anything flagged or marked high importance. Focused and Other come from Outlook's own verdict, the same one the Focused Inbox uses, so it is trained on how you actually treat your mail rather than guessed here. Notifications is the one judgement this makes on its own: senders whose address is a machine's, like no-reply, notifications, or a ticketing or build system. A person mailing from one of those domains still reads as a person. It sits above Other because ticket and build mail is work. Sections fold shut and remember it, and the split applies to the unified list, an inbox, and Unread Mail; browsing a filing folder or Sent Items stays flat, and a search always does.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailSplitInbox).onChange((v) => {
							s.mailSplitInbox = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Show as conversations",
				desc: "Group a back-and-forth into one row that expands, like Outlook's Show as Conversations. Search results are never grouped.",
				help: "A conversation collapses to a single row carrying everyone who wrote, the message count, and the marks off every message in it, so an attachment three replies down is still visible from the list. The twisty on the left opens it in place, and the reading pane lists the whole thread so you can walk it without going back. Ticking, archiving, or deleting a collapsed conversation takes all of it, which is what Outlook does. Search results stay ungrouped on purpose: a search asks about messages, and burying a match inside a collapsed thread would misreport what matched.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailConversations).onChange((v) => {
							s.mailConversations = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Sender photos",
				desc: "Show a sender's real profile picture in place of their initials. Needs a reconnect to grant the photo permission.",
				help: "Pictures come from your organization's directory, so colleagues have one and outside senders almost never do; anyone without falls back to the lettered circle. It needs the delegated ProfilePhoto.Read.All permission on the Azure app registration and a Reconnect on each account, and until then this quietly does nothing rather than erroring. Faces are fetched a few at a time as a list is drawn, cached, and kept in the disposable cache file so they are there the next time you open Obsidian.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailPhotos).onChange((v) => {
							s.mailPhotos = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Toolbars",
				desc: "Which buttons the mail and calendar toolbars carry, and in what order. Anything left off still has its keyboard shortcut, right-click entry, and palette line.",
				build: (st) => {
					st.addButton((b) =>
						b.setButtonText("Mail").onClick(() => {
							const leaf = this.app.workspace.getLeavesOfType(VIEW_TYPE_MAIL)[0]?.view;
							if (leaf instanceof MailView) leaf.openToolbarEditor();
							else new Notice("Power Desk: open the inbox first, so the toolbar can redraw as you change it.");
						})
					);
					st.addButton((b) =>
						b.setButtonText("Calendar").onClick(() => {
							const leaf = this.app.workspace.getLeavesOfType(VIEW_TYPE)[0]?.view;
							if (leaf instanceof PowerCalendarView) leaf.openToolbarEditor();
							else new Notice("Power Desk: open the calendar first, so the toolbar can redraw as you change it.");
						})
					);
				},
			},
			{
				name: "Announce new mail",
				desc: "A notice when mail arrives, and a desktop notification when Obsidian is not the window in front.",
				help: "Focused only follows Outlook's own verdict and stays quiet about the Other pile, which is usually what you want from a busy mailbox. Clicking a notice opens that message. Nothing is announced on the first check after Obsidian starts: the mail already in your inbox is not news, and mail that arrived while Obsidian was closed is not either. Read mail never announces itself, so a message you have already dealt with on your phone stays quiet here.",
				build: (st) => {
					st.addDropdown((d) => {
						d.addOption("off", "Never");
						d.addOption("focused", "Focused only");
						d.addOption("all", "Everything");
						d.setValue(s.mailNotify).onChange((v) => {
							s.mailNotify = v as "off" | "focused" | "all";
							save();
						});
					});
				},
			},
			{
				name: "Unread count on the ribbon",
				desc: "A red count on the mail icon down the side of the window.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailBadge).onChange((v) => {
							s.mailBadge = v;
							save();
							this.plugin.paintRibbonBadge();
						})
					);
				},
			},
			{
				name: "Check for new mail",
				desc: "How often an open mail view looks for new mail. Separate from the calendar's refresh, which is far slower on purpose.",
				help: "Mail syncs by delta, so a check that finds nothing is one small request per account and a minute is a comfortable cadence. Only what you are looking at is checked: the inbox lists, the folder tree's unread counts, and the folder currently open. Other folders you have visited refresh when you open them. The view also checks the moment you come back to the window or to the mail view, which is when stale mail is most obvious. This runs only while a mail view is open.",
				build: (st) => {
					st.addDropdown((d) => {
						d.addOption("0", "With the calendar");
						for (const [v, label] of [
							["30", "Every 30 seconds"],
							["60", "Every minute"],
							["120", "Every 2 minutes"],
							["300", "Every 5 minutes"],
						])
							d.addOption(v, label);
						d.setValue(String(s.mailPollSeconds)).onChange((v) => {
							s.mailPollSeconds = Number(v) || 0;
							save();
							this.plugin.mailCadenceChanged();
						});
					});
				},
			},
			{
				name: "Full width",
				desc: "The panel button in the mail and calendar headers folds the left sidebar away and brings it back. Nothing else changes your layout.",
				help: "The panel button at the left of the mail and calendar headers folds the whole left sidebar away, so mail or the calendar has everything except the ribbon down the far edge. Clicking anything on the ribbon brings the sidebar back, so the window is never a mode you cannot leave, and the 'Focus mode' command does the same thing if you would rather bind a hotkey. Nothing else moves the layout: opening mail or the calendar leaves the workspace exactly as you had it, because a window that rearranges itself as you change tabs is unsettling however sensible each rearrangement is.",
			},
			{
				name: "Categories",
				desc: "Make, recolor, and delete the categories your mailbox offers. Renaming is not something the API allows; see the note in the dialog.",
				help: "The mailbox's own category list, shared with Outlook and everything else that reads it. Make one, recolor one, delete one. There is no rename: Graph allows only the color to be changed once a category exists, so offering one would be offering a button that fails. Outlook appears to rename because it rewrites every message, event, and task carrying the old name at once. Deleting a category takes it off the list but leaves the label on messages that already have it.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Manage").onClick(() => new CategoriesModal(this.app, this.plugin, () => this.plugin.notify()).open()));
				},
			},
			{
				name: "Automatic replies",
				desc: "Your out-of-office, set here and sent by the mailbox. A band across the inbox reminds you while one is running.",
				help: "The out-of-office your mailbox sends, per account: on until you turn it off or between two times, with one message for colleagues and another for outsiders. The mailbox sends them, so they go whether or not Obsidian is open, and they are the same replies Outlook shows. While one is running, a band across the top of the inbox says so and offers to turn it off, which is the part that stops an away message answering for a fortnight after you got back.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Set up").onClick(() => new OutOfOfficeModal(this.app, this.plugin).open()));
				},
			},
			{
				name: "Inbox rules",
				desc: "Rules run in your mailbox, so they file mail while Obsidian is closed. These are the same rules Outlook shows.",
				help: "The same rules Outlook shows, edited here. They run in the mailbox rather than in Obsidian, so mail is filed whether or not this is open and on every device you read mail from. Power Desk offers a common subset of what a rule can do; a rule that also carries conditions or actions it cannot show says so, and keeps them untouched when you save.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Manage rules").onClick(() => new RulesModal(this.app, this.plugin).open()));
				},
			},
			{
				name: "Undo send",
				desc: "Seconds a sent message is held so you can take it back. The compose window closes right away; the message leaves when the time is up.",
				help: "Send closes the window immediately and a notice offers Undo for this long; taking it puts the window back exactly as it was, draft and all. The hold happens here rather than on the server, so a message still inside its window when Obsidian quits has not gone yet. Unloading the plugin or closing Obsidian normally lets anything waiting go, but a force quit cannot, which is why the choices are seconds rather than minutes. Off sends the instant you press Send.",
				build: (st) => {
					st.addDropdown((d) => {
						d.addOption("0", "Off");
						for (const n of [5, 8, 10, 20, 30]) d.addOption(String(n), `${n} seconds`);
						d.setValue(String(s.mailUndoSeconds)).onChange((v) => {
							s.mailUndoSeconds = Number(v) || 0;
							save();
						});
					});
				},
			},
			{
				name: "Snooze folder",
				desc: "The mailbox folder snoozed mail waits in, created on first use. It returns to the inbox when its time comes and Obsidian is running.",
				help: "Snoozed mail is moved into a real folder in your mailbox and moved back to the inbox when its time comes, so it is visible and recoverable from Outlook or any other client rather than hidden inside this plugin. The folder is created the first time you snooze something. Graph has no server-side snooze, so the return leg runs here: a message due while Obsidian is closed comes back the next time you open it.",
				build: (st) => {
					st.addText((t) =>
						t.setValue(s.mailSnoozeFolder).onChange((v) => {
							s.mailSnoozeFolder = v;
							save();
						})
					);
				},
			},
			{
				name: "Row density",
				desc: "Spacing in the message list. Compact fits the most on a screen; comfortable gives each row the most room.",
				help: "How much air each row in the message list gets. Compact fits the most on a screen, comfortable is the easiest to hit with a finger, cozy sits between them. This is spacing only; how many lines of preview text a row shows is the setting below.",
				build: (st) => {
					st.addDropdown((d) => {
						d.addOption("compact", "Compact");
						d.addOption("cozy", "Cozy");
						d.addOption("comfortable", "Comfortable");
						d.setValue(s.mailDensity).onChange((v) => {
							s.mailDensity = v as "compact" | "cozy" | "comfortable";
							save();
							this.plugin.notify();
						});
					});
				},
			},
			{
				name: "Message preview",
				desc: "Lines of body text under each message in the list. Off, the default, shows the sender and subject only, like Outlook Classic.",
				help: "How many lines of the message body sit under each row in the list, exactly Outlook's View > Message Preview. Off gives Outlook Classic's own two-line row, sender over subject, which fits about half again as many messages on a screen and is the fastest to scan once you know your mail.",
				build: (st) => {
					st.addDropdown((d) => {
						d.addOption("0", "Off");
						d.addOption("1", "1 line");
						d.addOption("2", "2 lines");
						d.addOption("3", "3 lines");
						d.setValue(String(s.mailPreviewLines)).onChange((v) => {
							s.mailPreviewLines = (Number(v) || 0) as 0 | 1 | 2 | 3;
							save();
							this.plugin.notify();
						});
					});
				},
			},
			{
				name: "Capture orders and bills",
				desc: "Scan incoming mail on the refresh timer and hand order confirmations and bills to Power Assistant, which turns them into notes you can report on. Needs Power Assistant with a transactions folder set; the rules live in its settings. Each matching message costs one AI call.",
				help: "Hands order confirmations and bills to Power Assistant on the refresh timer, which turns them into notes with the vendor, date, and line items filled in. Which messages qualify is decided by rules in Power Assistant's settings, and each match costs one AI call, so this stays off until you have set those rules up.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(this.plugin.settings.txnScan).onChange((v) => {
							this.plugin.settings.txnScan = v;
							save();
							if (v && !this.plugin.assistantTxn())
								new Notice("Power Desk: Power Assistant is not set up for transactions yet. Set a transactions folder in its settings.", 9000);
						})
					);
				},
			},
			{
				name: "Save attachments to",
				desc: "The folder outside the vault that an attachment's 'Save to folder' writes into; empty uses your Downloads folder.",
				help: "A folder outside the vault that an attachment's 'Save to folder' writes into. This is deliberately not a vault path: it is for files you want on the filesystem rather than in your notes. Empty uses your Downloads folder.",
				build: (st) => {
					st.addText((t) =>
						t.setValue(s.mailSaveFolder).onChange((v) => {
							s.mailSaveFolder = v;
							save();
						})
					);
				},
			},
			// One receipt per row. These were two unlabeled toggles sharing a row,
			// which needed the help popover to say which was which; now each says so
			// itself and, more to the point, each is findable by name in Obsidian's
			// own settings search, where "delivery receipt" used to match nothing.
			{
				name: "Ask for a read receipt by default",
				desc: "Asks the recipient's mail app to confirm they opened the message. Can still be set per message from Options in the compose window.",
				aliases: ["receipts"],
				help: "A standard mail field Outlook sets too, carried in the message headers: the recipient's own software decides whether to answer, and every serious one asks them first and lets them refuse. A receipt that never arrives therefore means nothing in particular, which is worth remembering before reading anything into silence. This ships off, because asking everyone you write to confirm they read you is a decision per message rather than a habit to acquire quietly.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailAskReadReceipt).onChange((v) => {
							s.mailAskReadReceipt = v;
							save();
						})
					);
				},
			},
			{
				name: "Ask for a delivery receipt by default",
				desc: "Asks the recipient's mail server to confirm the message arrived. Can still be set per message from Options in the compose window.",
				aliases: ["receipts"],
				help: "The companion to the read receipt above, and the same standard header field Outlook sets. The server at the other end decides whether to answer, so silence tells you nothing definite either way. It ships off for the same reason: asking for confirmation on everything you send is a decision per message rather than a habit to acquire quietly.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.mailAskDeliveryReceipt).onChange((v) => {
							s.mailAskDeliveryReceipt = v;
							save();
						})
					);
				},
			},
			{
				name: "Signatures",
				desc: "Named signatures, with a separate one for new messages and for replies, per account.",
				help: "As many signatures as you want, named, with each account choosing one for new messages and another for replies, since the full block belongs on a first message and rarely on the fourth reply of a thread. The editor is a real one: bold, italic, lists, links, and an image straight from the vault, so a logo goes in without hand-written HTML. A signature made before this existed was carried into the list automatically and is set on every account.",
				build: (st) => {
					st.addButton((b) => b.setButtonText("Edit signatures").onClick(() => new SignaturesModal(this.app, this.plugin).open()));
				},
			}
		);

		/* ---------------- Event notes ---------------- */

		const notes: Row[] = [
			intro("Every event can carry a note in your vault: frontmatter for querying, attendees as links so people pages connect, and the body all yours."),
			{
				name: "Notes folder",
				help: "Where an event's linked note is created. Every event can carry one note, and they all land here so the folder doubles as a record of what happened when.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder("Calendar").setValue(s.notesFolder).onChange((v) => {
							s.notesFolder = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Filename template",
				help: "Tokens: {{date}} (2026-07-17), {{time}} (09.30, empty for all-day), {{title}}, {{calendar}}. The result is sanitized for the filesystem, so a title's slashes or colons cannot break the path.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder("{{date}} {{title}}").setValue(s.noteNameTemplate).onChange((v) => {
							s.noteNameTemplate = v;
							save();
							this.plugin.notify();
						})
					);
				},
			},
			{
				name: "Open notes in a new tab",
				help: "Whether opening an event's note replaces the current tab or opens beside it. On for keeping the calendar visible while you write.",
				build: (st) => {
					st.addToggle((t) =>
						t.setValue(s.notesInNewTab).onChange((v) => {
							s.notesInNewTab = v;
							save();
						})
					);
				},
			},
			{
				name: "People folder",
				desc: "Where clicking an attendee lands. Empty borrows Power Assistant's People folder.",
				help: "Attendee and organizer names on an event card are links: click one to open that person's page, created on first visit. With Power Assistant installed these are the same pages its People hubs build on.",
				build: (st) => {
					st.addText((t) =>
						t.setPlaceholder(this.plugin.personFolderPath()).setValue(s.peopleFolder).onChange((v) => {
							s.peopleFolder = v;
							save();
						})
					);
				},
			},
		];

		return [
			{
				id: "microsoft",
				label: "Microsoft 365",
				groups: [
					{ heading: "Microsoft 365 accounts", rows: graphAccounts },
					{ heading: "Microsoft 365 app", rows: graphApp },
				],
			},
			{
				id: "google",
				label: "Google",
				groups: [
					{ heading: "Google accounts", rows: googleAccounts },
					{ heading: "Google app", rows: googleApp },
				],
			},
			{ id: "caldav", label: "CalDAV", groups: [{ heading: "CalDAV accounts", rows: caldav }] },
			{ id: "ics", label: "ICS feeds", groups: [{ heading: "ICS feeds", rows: ics }] },
			{ id: "vault", label: "Vault notes", groups: [{ heading: "Vault notes", rows: vault }] },
			{
				id: "calendar",
				label: "Calendar",
				groups: [
					{ heading: "Views", rows: views },
					{ heading: "Refresh", rows: refreshRows },
					{ heading: "Availability", rows: availability },
					{ heading: "Weather", rows: weather },
				],
			},
			{ id: "mail", label: "Mail", groups: [{ heading: "Mail", rows: mail }] },
			{ id: "notes", label: "Notes", groups: [{ heading: "Event notes", rows: notes }] },
		];
	}
}
