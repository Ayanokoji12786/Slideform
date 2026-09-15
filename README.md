# PowerPoint to Excel Converter for macOS

This app converts native PowerPoint table cells to Excel. It does not infer, calculate, or fill in missing values. Empty PowerPoint table cells remain empty in Excel.

## Run locally

```zsh
cd "/Users/MAC/Desktop/mom porject"
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
python3 main.py
```

Choose a `.pptx` file. The suggested Excel filename keeps the same base name, changing only `.pptx` to `.xlsx`.

When a PowerPoint file is selected, the app scans its native slide text and table rows for `Demand Priority`. It pre-fills every matching slide number. You can review or replace the selection with a comma-separated list such as `6` or `12, 13, 15`. Choose **All native tables on each selected slide** when a slide has more than one table that you want copied.

The original slide-image export remains available through **Insert slide images**. That option requires Microsoft PowerPoint for Mac and macOS Automation permission.

## Share with other Mac users

For a small team, zip this project folder and share it through your approved internal file-sharing system. Each user needs Python 3 and, only for slide-image export, Microsoft PowerPoint for Mac. They run the commands in **Run locally** once.

For a double-clickable app, package it on a Mac with the target processor architecture:

```zsh
source .venv/bin/activate
python3 -m pip install pyinstaller
pyinstaller --windowed --name "PowerPoint to Excel Converter" main.py
```

Distribute `dist/PowerPoint to Excel Converter.app`. For organization-wide distribution, sign and notarize the `.app` with your organization's Apple Developer certificate. Unsigned apps can show a Gatekeeper warning.

## macOS permissions

Table extraction reads the `.pptx` locally and does not require PowerPoint or Automation permission. Slide-image export can prompt for permission. Allow the launching app, usually Terminal, iTerm, or the packaged converter app, to control Microsoft PowerPoint at:

`System Settings → Privacy & Security → Automation`

## Offline browser app for any device

The `offline-app` folder contains a responsive installable web app. It has no backend and does not upload source files. Build its distributable files with:

```zsh
cd "/Users/MAC/Desktop/mom porject/offline-app"
npm install
npm run build
```

Share the resulting `offline-app/dist` folder through an approved static web host or package it as a Progressive Web App. Once opened and installed in a current browser, it works offline on macOS, Windows, Android, and iPadOS.

For native table extraction, upload the original `.pptx`. The app finds slides containing `Demand Priority` and downloads an `.xlsx` with the original table cells.

For full-slide images on any device, first export the presentation to PDF in PowerPoint or another presentation app. Then upload the PDF to the offline app. It renders each selected PDF page locally and places it as an image in its own Excel worksheet. This uses the PDF created by the presentation app as the visual source, which avoids a browser trying to reproduce PowerPoint-only effects.

New visitors see a short guided demo automatically the first time they open the app (a three-step walkthrough of choosing a file, reviewing slides, and downloading). It can be reopened anytime from **Guided demo** in the top bar, and is keyboard-navigable (arrow keys to move between steps, Escape to close).

### Preview before sharing

Check the built app in a real browser before distributing it:

```zsh
cd "/Users/MAC/Desktop/mom porject/offline-app/dist"
python3 -m http.server 8080
```

Open `http://localhost:8080` and try the full flow: pick a `.pptx`, confirm the auto-selected slides, and convert. Stop the server with `Ctrl+C` when done.

Access it here: https://slideform1.netlify.app/
