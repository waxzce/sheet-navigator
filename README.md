# Sheet Navigator

> A side panel for Excel that lists every sheet in your workbook. Open source, cross-platform, no account needed.
>
> Live demo & install: **[sheet-navigator.waxzce.org](https://sheet-navigator.waxzce.org/)** · 🇫🇷 [README en français](README.fr.md)

Cross-platform Office Add-in (Mac, Windows, Web) that adds a side panel to Excel listing all the sheets in the active workbook. Inspired by what Kutools does for sheet navigation, built on the modern Office.js platform, hosted on **Clever Cloud Static** and deployed via **GitHub Actions**.

## Features

- Compact list of all sheets, click to activate (selects `A1`)
- Colored dot on the left (`tabColor`), name in the middle, action icons on hover
- **Folder hierarchy via `/`** in sheet names: `Region/France` and `Region/Germany` appear as children of an expandable "Region" folder. Unlimited depth. Drag between folders renames automatically. Collapse state persisted per workbook.
- Accent-insensitive, case-insensitive search
- Toolbar: show/hide hidden sheets, sort by workbook or alpha, refresh
- Toggle individual sheet visibility via the eye icon
- Pin frequent sheets (pin icon on hover), persisted **per workbook** via `Office.context.document.settings`
- Drag-and-drop reorder
- Real-time sync via `onAdded`, `onDeleted`, `onActivated`, `onNameChanged` Office.js events

## Stack

- **Office.js** (classic XML manifest, not the unified JSON)
- **TypeScript** strict
- **React 18** + **Fluent UI v9** (`@fluentui/react-components`)
- **Webpack 5** (HTTPS dev server on `:3000`)
- **Clever Cloud Static** (static-apache) with platform-side build via `CC_PRE_BUILD_HOOK`

## Prerequisites

- macOS or Windows, Node.js LTS (≥ 20) — `brew install node` or `nvm install --lts`
- Recent Excel (Microsoft 365)
- [`clever-tools`](https://www.clever-cloud.com/doc/clever-tools/getting_started/) for deployment: `npm install -g clever-tools && clever login`

## Quick start

```bash
npm install
npm run dev              # webpack-dev-server HTTPS on :3000
npm run sideload:mac     # installs manifest.local.xml in Excel's wef folder
```

On first `npm run dev`, the Office template installs a **local HTTPS certificate** into your keychain (asks for your password) — required, because Office refuses HTTP except for `localhost`.

Then:
1. Fully quit Excel (`Cmd+Q` on Mac, close all windows on Windows)
2. Open a workbook
3. **Insert → Add-ins → My Add-ins → Developer Add-ins → Sheet Navigator**
4. The "Sheet Navigator" button appears in the **Home** ribbon tab

To open **Safari devtools** on Mac: right-click in the panel → *Inspect Element*.

## npm scripts

| Script | Action |
| --- | --- |
| `npm run dev` | Start webpack-dev-server HTTPS on `:3000`, hot reload |
| `npm run build` | Production build into `dist/`, hashed bundles |
| `npm run sideload:mac` | Copy `manifests/manifest.local.xml` into the wef folder |
| `npm run sideload:mac -- --prod` | Same but with `manifest.prod.xml` (prod URL) |
| `npm run sideload:mac:clean` | Remove every Sheet Navigator manifest from wef |
| `npm run deploy` | Deploy to Clever Static (`clever deploy --force`) |
| `npm run validate` | Validate the local manifest |

## Hosting on Clever Static

The Clever Cloud app is a **static-apache** configured to build the project on the platform side.

App env vars:

| Variable | Value |
| --- | --- |
| `CC_PRE_BUILD_HOOK` | `npm ci && npm run build` |
| `CC_WEBROOT` | `/dist` |
| `CC_COMPOSER_VERSION` | `2` |

The custom domain `sheet-navigator.waxzce.org` is attached to the app, with a Let's Encrypt cert auto-provisioned by Clever Cloud. The versioned `.clever.json` file references the `app_id`, so `clever deploy` works without prior `clever link`.

## Local deploy

```bash
clever login            # once
npm run deploy          # clever deploy --force
```

Production URL: `https://sheet-navigator.waxzce.org/`

## GitHub Actions deploy

The workflow `.github/workflows/deploy-clever.yml` runs on `push` to `main` or via `workflow_dispatch`.

Two secrets to configure in **Settings → Secrets and variables → Actions**:

| Secret | How to get it |
| --- | --- |
| `CLEVER_TOKEN` | `cat ~/.config/clever-cloud/clever-tools.json` after a local `clever login`, `token` field |
| `CLEVER_SECRET` | Same file, `secret` field |

## Distribute to teammates

Landing page: **https://sheet-navigator.waxzce.org/** with auto OS-detection + step-by-step install.

### Mac

```bash
curl -o ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/sheet-navigator.manifest.xml \
  https://sheet-navigator.waxzce.org/manifest.xml
```

Then restart Excel and go to **Insert → Add-ins → My Add-ins → Developer Add-ins**.

### Windows

Easiest: visit https://sheet-navigator.waxzce.org/, download `manifest.xml`, then in Excel **Insert → My Add-ins → Upload My Add-in** and pick the downloaded file.

Multi-user alternative: drop `manifest.xml` on a SMB share or synced OneDrive, then in Excel → File → Options → Trust Center → Settings → **Trusted Add-in Catalogs** → paste the UNC URL → check *Show in menu* → restart Excel.

### Microsoft 365 Centralized Deployment

If your org has an M365 admin tenant, an admin can push the manifest from **M365 admin center → Integrated apps → Upload custom apps**. The add-in then appears automatically in every targeted user's Excel — no manual sideload.

## Code structure

```
src/
  taskpane/
    index.tsx              # React + FluentProvider bootstrap
    App.tsx                # Root component, search + sort
    taskpane.html
    taskpane.css
    hooks/
      useWorksheets.ts     # Office.js wrapper: load + events + actions + pinning + folders
    components/
      SearchBar.tsx
      Toolbar.tsx
      SheetList.tsx        # Recursive tree rendering, drag-and-drop with rename
      SheetRow.tsx
      FolderRow.tsx        # Folder header with disclosure triangle
    utils/
      tree.ts              # Pure tree builder from "/"-delimited sheet names
  commands/
    commands.ts            # Function file referenced by the manifest (empty)
    commands.html
  index.html               # Marketing landing page (served at /)
  install.html             # Install instructions (Mac + Windows)
manifests/
  manifest.local.xml       # SourceLocation = https://localhost:3000/
  manifest.prod.xml        # SourceLocation = https://sheet-navigator.waxzce.org/
scripts/
  sideload-mac.mjs
.github/workflows/
  deploy-clever.yml
```

## Gotchas

- **HTTPS required**: Office refuses HTTP except for `localhost`. Clever Static + LE cert handles this natively.
- **Wildcard cert rejected by Office on Windows**: this is why we're on a custom domain `sheet-navigator.waxzce.org` (dedicated Let's Encrypt cert) rather than the wildcard `*.cleverapps.io` or `*.cellar-c2.services.clever-cloud.com`. Windows' Office stack does TLS pinning strictness that the Mac/Web clients don't.
- **Stubborn Office cache on Mac**: if a manifest change isn't picked up, run `npm run sideload:mac:clean`, then wipe the Office cache: `rm -rf ~/Library/Containers/com.Microsoft.OsfWebHost/Data/*`, then restart Excel.
- **Panel is always on the right**: Microsoft enforces a minimum width (~320 px) and doesn't allow docking on the left.
- **No unified JSON manifest yet**: the format is in preview on Mac, we stay on XML. Migration possible later via `npx office-addin-project convert`.

## Contributing

PRs welcome. The codebase is small (a handful of files) and the tree builder is pure — easy to extend.

If you want to add a feature:
1. Open an issue first describing the use case
2. Fork and branch
3. Run `npm run dev` for the live HTTPS panel + `npm run sideload:mac` to wire it into your Excel
4. Submit a PR

## License

MIT — see [LICENSE](LICENSE).
