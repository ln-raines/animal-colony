# Mouse Colony System — redesign

Replaces the old per-colony 3-tab setup with one shared `Animals` sheet,
a single combined `Cage_Chart`, automated weaning/breeder alerts, and an
external searchable dashboard (with the same email-approval reservation
flow) for your shared colonies.

Files:
- `Config.gs`, `Setup.gs`, `Triggers.gs`, `CageChart.gs`, `BreederWatchlist.gs`, `WebApp.gs`
  → paste into the Apps Script project bound to your spreadsheet.
- `docs/index.html`, `docs/app.js`, `docs/styles.css`
  → the dashboard frontend, deployed via GitHub Pages.

## 1. Set up the spreadsheet

You can do this in your existing spreadsheet or a fresh copy — your call.
If you reuse the existing one, this does **not** touch your current tabs;
it only adds the new ones listed below.

1. Open **Extensions → Apps Script** on your spreadsheet.
2. Delete the old `processEdit`/`syncToGroupSheet` code (or keep it in a
   separate unused file for reference — just make sure the old `onEdit`
   simple trigger function isn't also firing).
3. Create one script file per file above (File → New → Script) and paste
   in each one's contents, keeping the same filenames minus `.gs`
   (Config, Setup, Triggers, CageChart, BreederWatchlist, WebApp).
4. In the function dropdown at the top, select **setupAll** and click **Run**.
   Approve the permissions prompt (it needs access to this spreadsheet and
   to send email on your behalf for approvals/digests).
5. This creates: `Animals`, `Archive`, `Litters`, `Users`, `Colonies`,
   `Genotype_Colors`, `Cage_Chart`, `Breeder_Watchlist`. It's safe to
   re-run `setupAll` any time — it only repairs structure/validation, never
   touches your animal rows.

## 2. Migrate your data

For each of the 6 colonies:
1. Copy each animal's row from its old "line-by-line inventory" tab into
   `Animals`, filling in the `Colony` column (must match a name in the
   `Colonies` sheet — edit that sheet first if a colony's name differs).
2. Set `Cage Tag` using the `<Dorm>-<RowLetter><ColNumber>` convention,
   e.g. `1R-A3` (row letters A–I top to bottom, column numbers 1–7 left
   to right within a dorm).
3. For breeders, fill in `Breeder Paired Date` and set `Status` to `Breeder`.
4. Leave `Number of Litters` / `Date of Last Litter` blank — they're
   recomputed automatically from the `Litters` sheet.
5. For any historical litters you want reflected immediately, add rows to
   `Litters` (Colony, Breeding Cage Tag, Dam Tag *only if harem*, DOB).
   Wean date and breeder stats fill in automatically.
6. Add your teammates to `Users` (Name, Email) and your six colonies'
   rows in `Colonies` (already seeded by setup — just fix dorm/shared/
   approver-email columns as needed).
7. Add a hex color per genotype you use in `Genotype_Colors`.
8. Run **refreshCageChartAll** and **refreshBreederWatchlist** once by
   hand (function dropdown → Run) to populate them immediately — after
   that they keep themselves current on every edit, plus a 5am daily
   safety sweep.

## 3. Set the dashboard password

In the Apps Script editor, function dropdown → select **setDashboardPassword**,
click Run, but first temporarily edit the line at the bottom of the
function list to call it with your chosen password, e.g.:

```js
function runOnce() { setDashboardPassword('your-chosen-password'); }
```

Run `runOnce`, then delete that helper function. (Script Properties aren't
visible in the UI by default, which is why this goes through a function
call rather than a settings page.) Share the password with lab members
however you'd like (Slack, a sticky note — your call per the design doc).

## 4. Deploy the Web App backend

1. In Apps Script: **Deploy → New deployment**.
2. Type: **Web app**.
3. Execute as: **Me**. Who has access: **Anyone** (the password gate in
   `WebApp.gs` is what actually restricts use — "Anyone" here just means
   "anyone with the link can reach the script," required for a public
   frontend to call it).
4. Deploy, and copy the **Web app URL** (ends in `/exec`).
5. Open `docs/index.html` and replace `PASTE_YOUR_WEB_APP_EXEC_URL_HERE`
   with that URL.

Whenever you edit `WebApp.gs` (or any other file) afterward, you'll need
**Deploy → Manage deployments → Edit (pencil) → New version** for changes
to take effect — editing the code alone doesn't update a live deployment.

## 5. Publish the frontend on GitHub Pages

1. Create a new GitHub repo (can be private if you're on GitHub paid
   Enterprise with private Pages; otherwise Pages sites are public by
   default — see note below).
2. Push the contents of the `docs/` folder to that repo, either at the
   repo root or in a `docs/` folder.
3. In the repo: **Settings → Pages** → set source to the branch/folder
   you used.
4. GitHub gives you a URL like `https://yourusername.github.io/your-repo/`.
   Share that with lab members alongside the password.

**Note on privacy:** GitHub Pages sites are public to anyone with the
link unless your org is on GitHub Enterprise Cloud with private Pages
enabled. The password gate in the app is the actual access control;
treat the URL as "unlisted," not secret.

## 6. Ongoing use

- **Adding/removing dropdown names** (Users, Colonies): just edit those
  sheets directly — everything reads live, no script edits needed.
- **Adding a new shared colony**: set its `Shared?` to `Y` and fill in
  `Approver Email` in `Colonies`. It shows up in the dashboard's colony
  filter and search automatically — no code change.
- **Weaning**: enter a litter's DOB in `Litters`; wean date (DOB + 21
  days) and the Active→Weaned flip are automatic. Change the default by
  editing `WEAN_DAYS_DEFAULT` in `Config.gs`.
- **Breeder alerts**: `Breeder_Watchlist` recomputes on every relevant
  edit and nightly; a digest emails everyone in `Users` every Monday 8am.
  Thresholds (90 days / 65 weeks) are `BREEDER_STALL_DAYS` /
  `BREEDER_MAX_AGE_WEEKS` in `Config.gs`.
- **Cage chart**: fully derived from `Animals`/`Litters` — don't hand-edit
  animal slots, they'll be overwritten on the next refresh. Hand-typed
  notes in the notes row are preserved unless there's an active/recent
  litter to report.
- **Harem breeding**: give the litter row a `Dam Tag` matching the
  specific female's `Physical Tag` so her stats (not the whole cage's)
  get credited.

## What changed from the old system, and why

- One `Animals` sheet instead of 6 colonies × 3 tabs — same reservation
  states (Available/Requested/Reserved), same color coding, generalized
  to work across every colony with a `Colony` column.
- No more `syncToGroupSheet()` / mirrored "Group Sheet" — that's what
  caused the data-validation copy bugs you ran into. The new dashboard
  reads live from this one sheet via the Web App, so there's nothing to
  keep in sync.
- Cage chart zones are **fully occupancy-driven**: a cage's colony/color
  is whatever `Animals.Colony` currently occupies it. Colonies can grow
  or shrink into any cage without touching config.
