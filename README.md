# Power Desk

Your calendars and mail inside Obsidian: Microsoft 365 (several accounts at once), Google Calendar, CalDAV (iCloud, Fastmail, Nextcloud), and ICS feeds in month, week, day, and agenda views, a triage inbox beside them, and a linked note for every event.

![A week view of a calendar with a recurring morning standup Monday through Friday, several timed meetings, a month picker and an agenda rail down the left side, and view switchers for month, week, work week, day and agenda](docs/images/week-view.png)

Week view with the month picker and an agenda rail beside it. The source here is **Vault notes**, which turns any note with a date property into an event, so a vault can
fill a calendar with no account connected at all. Month, work week, day, and agenda sit
behind the switcher, and Filter narrows by calendar.

Formerly named Power Calendar; the plugin id changed to `powerdesk` with the rename, and settings migrate automatically on the first deploy. The ```` ```power-calendar ```` code block keeps working forever (```` ```power-desk ```` is its new alias), so existing daily-note templates are untouched.

Microsoft 365 and Google events are fully editable: create them by dragging across empty grid, move and resize them with the mouse, edit and delete from the event card. CalDAV and ICS stay read-only. Notes land in your vault, one per event, on your terms.

## Views

- **Month**: a classic grid. All-day and multi-day events render as banners across the week rows; timed events are chips inside each day. A crowded day collapses to "+N more", which jumps to that day.
- **Week**, **Work week** (Monday through Friday), and **Day**: a timed grid with side-by-side layout for overlapping events, an all-day strip on top, and a red line that knows what time it is.
- **Agenda**: the next N days as a readable list, with one-click join buttons for meetings.

Click any event for its card: time, calendar, location, organizer, attendees, description, RSVP buttons on invites, and the actions (join the meeting, open or create its note, edit, delete, open the original). When Power Assistant is installed, the card also offers "Capture meeting", which opens its New meeting dialog prefilled with the event's title, time, attendees, and Teams link.

Click the title in the header for a **mini month navigator**: density dots per day, arrows across months, one click lands anywhere. A **second time zone** (settings) draws its clock beside the hours in the week and day grids. And **Copy free slots** (command palette) reads your next five workdays and puts your open times on the clipboard as paste-able text, the serverless cousin of a scheduling link.

## Creating and editing events (Microsoft 365 and Google)

- Drag across empty space in the week or day grid to draft an event there; double-click for a half-hour slot; double-click a month cell for an all-day event, or drag across month cells for a multi-day one; press `C` or use the + button for a quick event at the next half hour.
- Drag an event block to move it (across days too); drag its bottom edge to resize. Escape cancels a drag. Times snap to 15 minutes.
- Edit and Delete live on the event card. Deleting a meeting you organize notifies attendees; deleting one occurrence of a series cancels just that occurrence. Deleted events land in your mailbox's Deleted Items, so they are recoverable.
- Microsoft 365 events are editable when you organize them and the account is connected with edit access; Google events are editable in any calendar where your account can write. Events from CalDAV and ICS sources never offer editing.
- With several writable calendars connected, the New event dialog offers a calendar picker; quick-create targets your default (Microsoft) or primary (Google) calendar.
- **Invite people** from the New event and Edit dialogs: email addresses, comma separated ("Bob <bob@x.com>" works too). Invitations and updates go out when you save; an untouched invite list is never resent.
- **Repeat** when creating: daily, every weekday, weekly, monthly, or yearly. Editing or deleting a repeating event asks whether you mean this occurrence or the whole series, exactly like Outlook; dragging one occurrence only moves that occurrence.
- **RSVP** right on the event card: invites you receive show Yes / Maybe / No, with your current answer highlighted; the organizer gets your response like any other reply.
- **Show as** Busy, Free, or Tentative when creating or editing; free events also stop blocking your Copy free slots availability.
- **Check availability** while drafting a Microsoft 365 meeting: free/busy bars for each invitee (same organization), your draft slot overlaid in red, and a per-person free / tentative / busy verdict that follows as you move the times. People the tenant gives you no visibility into show as such instead of failing the panel.
- **Email attendees** from any event card: a small compose with everyone addressed and the event as the subject, sent from your own mailbox (or handed to your mail app).

## Mail

Calendar and mail share the same accounts, so Power Desk carries a **triage inbox**: recent messages across every mail-enabled Microsoft account in one list (unread bold, an account color dot per row), a text-first reading pane, reply, archive, and compose. Two bridges make mail belong here:

- **Make event**: an email becomes a prefilled New event dialog, with the sender already invited and the reply prefixes stripped from the title.
- **Save to note**: the message lands in your vault as a dated note with the text and a link back to the original.

A **folder pane** on the left shows each account's full folder tree with unread counts, your filed subfolders included, exactly as in Outlook. Folders with children carry a collapse chevron; a collapsed branch rolls its whole subtree's unread count up onto itself, and the folds are remembered. The account header folds too, taking the whole mailbox with it and wearing the account's unread total while closed. Above each tree sits **Unread Mail**, the Outlook-style search folder: every unread message across the inbox and all its subfolders in one list, each tagged with the folder it lives in, and (with the keep-unread option on) nothing marks itself read while you walk it. Click any folder to browse it; its first message opens in the reading pane right away (without marking anything read on its own), "All inboxes" on top returns to the unified list, and the panel button in the header collapses the pane. On phones and narrow panes the view drills one screen at a time instead, exactly like the phone Mail app: Mailboxes, then the message list, then the message, with back in the header. Hover any message for quick actions right in the list: reply, mark read or unread, archive; the reading pane carries the same plus Mark unread after a message auto-marks itself read. An **Unread filter** in the header shows only unread mail (with a live count), keeping the message you are reading visible until you move on. **Mark-as-read behavior** is yours to pick in settings, exactly as in Outlook: on selection, a few seconds after selection, when the selection changes, or only by hand; and while the Unread filter is on, an option keeps everything unread unless you explicitly mark it. **Message preview** in settings sets how many lines of body text sit under each row, the same Off / 1 / 2 / 3 that Outlook's View menu offers; Off gives Outlook Classic's own two-line row, sender over subject in the accent color, which fits about half again as many messages on a screen.

**Attachments** ride a bar under the message actions, each on a chip with a colored file-type tile: PDF red, Word blue, Excel green, PowerPoint orange, images purple, archives amber, and so on down to a neutral gray for anything unrecognized, so a bar of files is scannable by color before a single name is read. The chip carries the name, the size, and a chevron that opens its actions: **Preview** opens the file in an Obsidian tab (PDFs, images, audio, and video all render there), **Open** hands it to the operating system's default app, **Save to vault** and **Save to folder** write it where you want it, and **Copy name** takes the filename. A message with several attachments gets a **Save all** button at the head of the bar. Open the view from the mail ribbon icon or the "Open inbox" command. The **search button** in the header searches the whole mailbox of every mail account at once, Outlook style (subject, sender, and body), results tagged with their folder and newest first. It is deliberately a triage surface, not a mail client: no folder management and no rules; "Open in Outlook" is one click for anything deeper. Mail arrives per account after a one-time Reconnect that grants the mail permission (the Azure app registration needs delegated `Mail.Read` and `Mail.Send` added). Each account row in settings has an inbox toggle, so a calendar-only account stays calendar-only.
- **Duplicate** any event from its card into a prefilled New event dialog.
- **Meeting reminders**: while Obsidian is open, a card slides in before each timed meeting (lead time in settings, 0 turns it off): the meeting, how long you have, the time, and where, with a one-click **Join**. It waits to be dismissed rather than timing out, and **Snooze** puts it back in five minutes (or at the start time, whichever comes first). "Where" names the service (Zoom, Teams, Google Meet, Webex) instead of printing the join link, since the link is what the button is for, and a room name shows as itself. Declined meetings and all-day events stay quiet. **Show a test meeting reminder** (command palette) brings the card up on demand, built from your next meeting today at your configured lead time, so you can see what it will look like without waiting for one.
- **Week numbers** (settings): ISO week labels in the month rows and the week header.
- Editing never rewrites an event's description: the plugin only fetches a preview of the body, and writing that back would truncate the real one.

## Keyboard

The view is keyboard-first. With the calendar focused:

| Key | Action |
| --- | --- |
| `T` | Today |
| `M` `W` `D` `A` | Month, Week, Day, Agenda |
| `5` | Work week (Monday through Friday) |
| Left / Right | Previous / next period |
| `C` | New event at the next half hour |
| `/` | Find event |
| `R` | Refresh now |
| `Esc` | Close the event card, cancel a drag |

## Your agenda inside any note

Put a `power-calendar` code block in a note and the day's events render right there, Live Preview safe:

````
```power-calendar
```
````

With no options, the block reads the day from the note's own filename (a daily note named `2026-07-18 Saturday` shows that Saturday), falling back to today. Options, one per line: `date: 2026-07-20` pins a day, `days: 3` widens the span. Rows show time, title, location, and a join button; clicking one opens the calendar on that day. Drop the block into your daily-note template once and every day has its agenda.

There is also **Find event** (`/` in the calendar, or the command palette anywhere): fuzzy search over the coming two months of events by title, day, calendar, or attendee; Enter lands the calendar on its day.

## People and bases

Attendee and organizer names on an event card are links: click one to open that person's page (created on first visit) in your People folder, which borrows Power Assistant's when it is installed, so meetings connect to the same person hubs its captures build. And **Create events base** (command palette) drops a ready-made Bases table over your event-notes folder, using Power Bases' table when present.

On phones, the calendar opens in the view you pick under "Default view on phones" (agenda by default, which reads best on a narrow screen), and chips, rows, and buttons grow to finger size.

## Notes per event

Every event can carry a note. "Create note" on the event card writes one with queryable frontmatter (date, calendar, location, organizer) and the attendees as `[[links]]`, so people pages connect on their own. Events whose note already exists show a small dot. The folder and filename template are settings; templates know `{{date}}`, `{{time}}`, `{{title}}`, and `{{calendar}}`.

## Sources

Settings has a tab per source (Microsoft 365, Google, CalDAV, ICS feeds, and Vault notes), each with its own "Add account" button. Each account row expands to its calendars (toggles and colors) and can carry a friendly name; each calendar gets its own status chip in the footer, and everything refreshes on the cadence you pick. Several accounts of the same kind can be signed in side by side (work plus personal Microsoft accounts, more than one Google account).

### Microsoft 365

Your Outlook calendars through **your own** Azure app registration, with full event editing. Sign-in uses the device-code flow: the plugin shows a code, you approve it in your own browser, and no password ever touches Obsidian. This flow also works on phones and tablets, where browser-redirect sign-ins cannot.

Setup, once: press "Show the steps" in settings (Microsoft 365 tab) for a walkthrough of the Azure portal with the exact value for every field. In short:

1. In your browser, open portal.azure.com signed in with the account that should own the app, then App registrations, New registration. Supported account types: "Any Entra ID Tenant + Personal Microsoft accounts" (older portals word it as "accounts in any organizational directory and personal Microsoft accounts"); Redirect URI: empty.
2. Under Authentication, turn on "Allow public client flows".
3. Paste the Application (client) ID from the app's Overview page into the plugin's settings and press Connect. Single-organization apps also need the Directory (tenant) ID in the Tenant field.

You consent to the calendar and mail permissions at first sign-in; only organizations that block user consent need an admin to grant the delegated `Calendars.ReadWrite`, `Mail.Read`, and `Mail.Send` Microsoft Graph permissions under API permissions.

If Power Assistant is installed and already connected to Microsoft 365, Power Desk borrows its app registration automatically and step 1 through 3 disappear; you still sign in once, and your calendars appear.

A connection made by an older version of this plugin was consented read-only. It keeps working exactly as before; press Reconnect in settings once to grant edit access.

After connecting, every calendar in the account is listed in settings with its own toggle and color.

Personal outlook.com accounts: an app registered as "this organizational directory only" cannot sign them in. Press Add account and pick Personal: the account signs in through its own app registration, which the wizard creates with you on the spot (owned by the personal account) or reuses silently from an earlier personal account. The app is remembered on the account row, so reconnects use it automatically. Alternatively, widening the shared app's supported account types to include personal accounts makes the separate app unnecessary.

### Google Calendar

Google requires **your own** (free) Google Cloud project, because Google's terms forbid shipping shared credentials inside an open-source plugin. One-time setup, about five minutes:

1. At console.cloud.google.com, create a project.
2. APIs & Services, Library: enable the **Google Calendar API**.
3. APIs & Services, OAuth consent screen: choose **External**, fill the required fields, and press **Publish app**. Do not leave the project in Testing mode: test-mode refresh tokens expire every 7 days, which reads as being signed out weekly. Publishing an app you never submit for verification simply means a one-time "Google hasn't verified this app" warning at sign-in (Advanced, then continue).
4. APIs & Services, Credentials: create an **OAuth client ID** of type **Desktop app**. Paste its Client ID and Client secret into Power Desk's settings under "Google app".
5. Press Add account, pick Google, and finish the sign-in in your browser.

Sign-in itself runs on desktop (the browser hands the code back to Obsidian on a local port); after that the connection syncs to phones and tablets with the vault, where reading and editing work normally.

### Vault notes

The source no other calendar can have: your own notes. Add a vault source (a folder filter plus the frontmatter property holding the date) and every matching note renders as an event beside your real calendars. A bare date (`2026-07-17`) is all-day; a datetime (`2026-07-17T09:30`) is timed; an optional end property stretches either. Click the event to open the note. Drag a timed one and its frontmatter date is rewritten, which makes deadlines reschedulable from the calendar. Everything is local: no network, updates appear as you edit notes.

### CalDAV

iCloud, Fastmail, Nextcloud, Radicale, and anything else speaking CalDAV. Add an account with its server URL and credentials, and the plugin discovers the calendars and lists them with toggles.

- **iCloud**: server `caldav.icloud.com`, your Apple ID as username, and an app-specific password from appleid.apple.com (your normal password will not work).
- **Fastmail**: server `caldav.fastmail.com` with an app password.
- **Nextcloud**: `https://your.server/remote.php/dav` with your normal or app credentials.

### ICS feeds

Read-only iCalendar subscriptions: a published Outlook or Google calendar, the team schedule, holidays. Paste the URL (`webcal://` links work) and name it.

## Network use

This plugin talks to the network only for the calendars you configure:

- `login.microsoftonline.com` and `graph.microsoft.com` for Microsoft 365 sign-in, calendar access, and mail (inbox reads, replies, and sends you make yourself), only after you connect an account.
- `accounts.google.com`, `oauth2.googleapis.com`, and `www.googleapis.com` for Google sign-in and calendar access, only after you connect an account through your own Google Cloud project.
- The CalDAV servers you add, with the credentials you provide.
- The ICS feed URLs you add.
- `geocoding-api.open-meteo.com` and `api.open-meteo.com` for the optional weather strip, and only once you set a location for it: the place name you search for goes to the first, and the resulting coordinates to the second. Leave the weather location empty and neither is ever called. Open-Meteo needs no account and no key.

Fetched calendars, inboxes, and folder trees are cached to a `cache.json` beside the plugin, so both views paint instantly from the last known state on every launch and refresh quietly in the background; the file holds data only (never tokens) and can be deleted freely. Mail folders sync by Microsoft Graph delta token: after the first load of a folder, each refresh transfers only the messages that arrived, changed, or were removed, and the plugin quietly pre-syncs your inbox subfolders and the bodies of the messages at the top of the list. Folder clicks and message opens then come from the local replica, the way a desktop mail client in cached mode works. Calendars fetch the same way: each source loads a window reaching about six weeks back and three months forward, so moving between weeks and months shows events already on hand and the network only refreshes them in the background.

Nothing beyond the hosts listed above is contacted, and no telemetry of any kind exists. The only calendar writes are the ones you make yourself: creating, moving, editing, or deleting a Microsoft 365 event sends exactly that change and nothing more. Sign-in tokens and CalDAV credentials are stored in this plugin's `data.json` inside your vault; if you sync your vault, they sync with it.

## Install

Until a community-store release, install via BRAT or by copying `manifest.json`, `main.js`, and `styles.css` into `.obsidian/plugins/powerdesk/`. Developers: `npm install`, then `npm run dev` to build, `npm test` for the suite, `npm run deploy` to copy the build into every local vault.

Power Desk stands alone. It has no required companions; when siblings from the Power family are present (Power Assistant's Microsoft 365 app registration), it borrows quietly and works the same either way.

## License

MIT

## Support

Power Desk is built and maintained by one person. If it earns a place in your
daily vault, you can [buy me a coffee](https://buymeacoffee.com/powerplugins).
Nothing in the plugin is held back either way.
