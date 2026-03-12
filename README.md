# Google Docs Photo Report Automation

This repo contains a Google Apps Script for assembling photo-heavy engineering reports in Google Docs.

## Overview

The script is meant to solve a common reporting problem:

- engineers, inspectors, and field staff collect a folder of numbered site photos
- report captions are written in a Google Doc as numbered lines such as `12 - Roof flashing at east elevation`
- manually matching each caption to the correct photo is repetitive, slow, and easy to get wrong

Photo Report automates that workflow by scanning the document for numbered caption lines, matching each caption number to the corresponding image number in a Google Drive folder, and inserting the image directly above the caption.

The script is designed to make repeated runs safe:

- Photos live in one Google Drive staging folder and are named `Image 1`, `Image 2`, `Image 3`, and so on.
- Captions are typed in a Google Doc as plain text lines such as `1 - Front view of house`.
- Running the script inserts the matching image above each caption.
- Re-running the script is safe because inserted images are tracked with markers in their alt text and skipped on later passes.

## Features

- Visual Folder Selection: choose the Drive photo folder from a Google Picker dialog instead of pasting IDs or URLs.
- Time-Guarded Batching: long runs stop early with a clear warning so you can run `Insert Missing Images` again before Apps Script times out.
- Idempotent Image Markers: inserted photos are tagged in alt text so reruns skip work that is already complete.
- Deterministic File Matching: duplicate image numbers are resolved consistently and reported back in the summary.

## What it does

The script adds a `Photo Report` menu to a Google Doc with these actions:

- `Set Image Folder`: saves the Drive folder that holds the staged photos.
- `Insert Missing Images`: inserts only images that are not already above a matching caption.
- `Rebuild Images`: validates the caption numbering, removes previously inserted script-managed images, and pauses safely if the Apps Script time budget is exhausted before rebuilding the photo/caption sequence.

The script resizes wide images to fit the page and centers them.

## Visual Folder Selection

The `Set Image Folder` action opens a Google Picker dialog inside Google Docs so users can browse Drive folders visually instead of copying a folder ID by hand.

To use the picker in your own Apps Script project:

1. Attach the script to a standard Google Cloud project.
2. Enable the Google Picker API for that Cloud project.
3. Create a Picker API key restricted to `*.google.com` and `*.googleusercontent.com` referrers.
4. Update `PHOTO_REPORT_CONFIG.pickerDeveloperKey` and `PHOTO_REPORT_CONFIG.pickerCloudProjectNumber` in [Code.gs](./Code.gs).

## Important input rules

- Type captions as plain text, not as a Google Docs auto-numbered list.
- Use one caption per paragraph in the main body of the document.
- Format each caption like `1 - Caption text`, `2: Caption text`, or `3) Caption text`.
- Keep the image numbers aligned with the caption numbers.
- Use a dedicated staging folder for each report run so the numbering stays clean and predictable.

Why plain text instead of an auto-numbered list? Apps Script can reliably read the text `1 - Front view of house`, but it does not reliably expose the rendered list number from Google Docs numbering.

## Setup

1. Create or open the Google Doc you use for reports.
2. Open `Extensions` > `Apps Script`.
3. Replace the default script contents with the contents of [Code.gs](./Code.gs).
4. Replace the manifest with the contents of [appsscript.json](./appsscript.json).
5. Save the Apps Script project.
6. Reload the Google Doc.
7. Run `Photo Report` > `Set Image Folder` and choose the Drive folder that holds the project photo staging images.
8. Author captions in the document as plain text lines such as:

```text
1 - Front view of house
2 - Roof overview
3 - Rear elevation
```

9. Put the project photos into the configured Drive folder with names like:

```text
Image 1.jpg
Image 2.jpg
Image 3.jpg
```

10. Run `Photo Report` > `Insert Missing Images`.

## Recommended workflow

For repeated projects, the cleanest setup is:

1. Make one Google Doc template that already has this script attached.
2. Duplicate that doc for each new report.
3. Drop the new project photos into the staging folder.
4. Paste or type the numbered captions.
5. Run `Insert Missing Images`.
6. If the script pauses because of the Apps Script execution limit, run `Insert Missing Images` again to continue where it left off.

## Matching behavior

- The script looks for the last number in each image filename.
- `Image 12.jpg` matches caption `12 - ...`.
- If the same caption number appears more than once in the document, the script stops before making changes and tells you which numbers must be fixed.
- If multiple files map to the same number, the script keeps one match and reports the duplicate number in the summary.
- If a caption number has no matching image, that number is listed in the summary.
- If a long run approaches the Apps Script time limit, the script pauses safely and the next `Insert Missing Images` run skips already managed images.

## Limits and notes

- Captions inside tables, headers, footers, or footnotes are not scanned.
- The script is designed for Google Drive, not Dropbox.
- Google Apps Script executions are capped at 6 minutes, so large reports may need more than one run.

## Future extensions

If you want to grow this later, the next logical additions would be:

- reading captions from a Google Sheet instead of the document body
- creating a title page or section breaks automatically
- storing per-document settings such as image width and validation preferences
