# Google Docs Photo Report

[![License](https://img.shields.io/github/license/tweakyourpc/gdocs-photo-report)](./LICENSE)
[![Lint](https://img.shields.io/github/actions/workflow/status/tweakyourpc/gdocs-photo-report/lint.yml?branch=main&label=lint)](https://github.com/tweakyourpc/gdocs-photo-report/actions/workflows/lint.yml)
[![Version](https://img.shields.io/github/v/release/tweakyourpc/gdocs-photo-report)](https://github.com/tweakyourpc/gdocs-photo-report/releases)

Google Apps Script for assembling photo-heavy engineering reports in Google Docs.

## Overview

This project solves a simple but expensive reporting problem:

- site photos are collected in Google Drive and named with report numbers such as `Image 12.jpg`
- captions are typed in a Google Doc as lines such as `12 - Roof flashing at east elevation`
- manually pairing each caption to the right image is repetitive, slow, and easy to get wrong

Photo Report scans the document, matches caption numbers to image numbers, inserts the image above the matching caption, and keeps the workflow safe to rerun.

## Features

- Live Progress Sidebar: run the tool from a persistent Google Docs sidebar with visible progress, counters, status messages, and batch resume controls.
- Visual Folder Selection: launch the Google Picker API from the sidebar to browse Drive folders visually.
- Manual Folder Fallback: if Picker credentials are not configured, the same sidebar still accepts a Drive folder URL or raw folder ID.
- Time-Guarded Batching: large reports are processed in resumable batches so the script stops cleanly before Apps Script execution limits become destructive.
- Auto Resume Flow: normal batches continue automatically, and time-limit pauses expose a resume button without sending the user back to the menu.
- Idempotent Image Markers: inserted images are tracked through alt text markers so repeated runs skip work that is already complete.
- Deterministic Matching: duplicate image-number collisions are resolved consistently and reported.
- Diagnostic Logs: if a run fails, the script creates a Google Doc log with the relevant caption numbers, file IDs, and error details.

## How It Works

1. Open the sidebar from the `Photo Report` menu in Google Docs.
2. Choose a Drive folder that contains numbered project photos.
3. Write plain-text caption lines in the body of the document, such as:

```text
1 - Front view of house
2 - Roof overview
3 - Rear elevation
```

4. Run `Insert Missing Images` or `Rebuild Images` from the sidebar.
5. Watch progress in the sidebar while the script advances through batches.
6. If Apps Script nears the execution limit, the sidebar pauses safely and resumes the next batch from the same UI.

## Sidebar Workflow

The sidebar is the main control surface for the project.

- `Open Sidebar` opens the persistent control center.
- `Set Image Folder` opens the same sidebar focused on folder selection.
- `Insert Missing Images` opens the sidebar and starts an insert run.
- `Rebuild Images` opens the sidebar and starts a rebuild run.

The progress panel shows:

- processed captions versus total captions
- inserted, removed, and remaining counts
- duplicate folder-number warnings
- missing image-number warnings
- diagnostic log links for failed runs

## Visual Folder Selection

The sidebar supports two folder-selection modes.

### Google Picker Mode

Use this when you want a visual Drive folder browser.

1. Attach the Apps Script project to a standard Google Cloud project.
2. Enable the Google Picker API for that Cloud project.
3. Create a Picker API key restricted to `*.google.com` and `*.googleusercontent.com`.
4. Open the sidebar and save the Picker API key plus Cloud project number in the Picker Settings panel.

Picker credentials are stored in `UserProperties`, which means one engineer can reuse the same credentials across multiple report documents without committing secrets into the repo.

If you want repo-level defaults for your own hosted copy, you can also set:

- `PHOTO_REPORT_CONFIG.defaultPickerDeveloperKey`
- `PHOTO_REPORT_CONFIG.defaultPickerCloudProjectNumber`

in [Code.gs](Code.gs).

### Manual Folder Mode

If Picker is not configured, the sidebar still works.

- Paste a Drive folder URL, such as `https://drive.google.com/drive/folders/<id>`
- Or paste a raw folder ID directly

The script validates and stores the folder for the current document.

## Important Input Rules

- Use plain text captions, not Google Docs auto-numbered lists.
- Use one caption per paragraph in the main document body.
- Format captions like `1 - Caption text`, `2: Caption text`, or `3) Caption text`.
- Keep image numbers aligned with caption numbers.
- Use a dedicated staging folder for each report run so numbering stays predictable.

Apps Script can read the literal text `1 - Front view of house`, but it does not reliably expose rendered list numbering from Google Docs auto-numbered lists.

## Setup

1. Create or open the Google Doc you use for reports.
2. Open `Extensions` > `Apps Script`.
3. Replace the default script contents with [Code.gs](Code.gs).
4. Add [Sidebar.html](Sidebar.html) as an HTML file in the Apps Script editor.
5. Replace the manifest with [appsscript.json](appsscript.json).
6. Save the Apps Script project and reload the Google Doc.
7. Open `Photo Report` > `Open Sidebar`.
8. Save a folder with either the Picker UI or the manual folder field.
9. Start `Insert Missing Images` or `Rebuild Images` from the sidebar.

## Matching Behavior

- The script looks for the last number in each image filename.
- `Image 12.jpg` matches caption `12 - ...`.
- If the same caption number appears more than once in the document, the run stops before making changes.
- If multiple files map to the same number, the script keeps one deterministic match and reports the duplicate number.
- If a caption number has no matching image, that number is listed in the progress summary.
- Re-running `Insert Missing Images` is safe because already-managed images are recognized and skipped.

## Diagnostic Logs

When a run fails, the script creates a temporary Google Doc that records:

- the report action that failed
- the source report document
- the configured Drive folder
- any caption numbers that were involved
- the exact Drive file IDs and file names tied to insertion errors

The sidebar links directly to the diagnostic log so the operator can inspect or share it.

## Development

The repo includes:

- [eslint.config.mjs](eslint.config.mjs) with Google JavaScript style enforcement for `Code.gs`
- [lint.yml](.github/workflows/lint.yml) for GitHub Actions linting on pushes and pull requests
- [clasp-push.yml](.github/workflows/clasp-push.yml) for production deployment on `main`
- GitHub issue templates for bug reports and feature requests
- an OpenGraph art brief in [docs/og-image-brief.md](docs/og-image-brief.md)

### Local Tooling

```bash
npm install
npm test
```

### GitHub Actions Deployment

The production `clasp` workflow expects these GitHub secrets:

- `CLASP_CREDENTIALS_JSON`
- `CLASP_SCRIPT_ID`

The workflow writes an ephemeral `.clasp.json` during CI and pushes the Apps Script project only from `main`.

## Repository Layout

- [Code.gs](Code.gs): Apps Script runtime, batching engine, diagnostics, and Drive matching logic
- [Sidebar.html](Sidebar.html): sidebar UI, progress updates, picker launcher, and resume controls
- [appsscript.json](appsscript.json): Apps Script manifest and OAuth scopes
- [LICENSE](LICENSE): MIT license

## Limits and Notes

- Captions inside tables, headers, footers, and footnotes are not scanned.
- The script is designed for Google Drive storage.
- Apps Script executions are capped at 6 minutes, so large reports are intentionally split into resumable batches.
