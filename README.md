# Duty Report Generator
### RGU · CSE — Unified System

A self-contained, browser-based tool for generating, editing, and exporting monthly teacher duty reports for Rajiv Gandhi University (CSE Department). No installation or server required — open the `.html` file in any modern browser and it works offline.

---

## Table of Contents

1. [Overview](#overview)
2. [Getting Started](#getting-started)
3. [Tab-by-Tab Guide](#tab-by-tab-guide)
   - [Setup](#1-setup-tab)
   - [Generate & Edit](#2-generate--edit-tab)
   - [Project Data](#3-project-data-tab)
4. [Exporting Reports](#exporting-reports)
   - [Print / Save as PDF](#print--save-as-pdf)
   - [Download as .docx](#download-as-docx)
5. [Data Persistence](#data-persistence)
6. [Fixing the .docx Report Layout](#fixing-the-docx-report-layout)
7. [Tips & Troubleshooting](#tips--troubleshooting)

---

## Overview

The tool generates a formatted monthly duty report by cross-referencing your weekly timetable against a selected month and year. It automatically skips Sundays, Saturdays, and any holidays or leave dates you have marked. The result is a two-column attendance table (left and right columns per page) ready to print, save as PDF, or download as a Word document.

---

## Getting Started

1. Open `[url](https://mainong-jenbum.github.io/dutyreport/` in any modern browser (Chrome, Edge, or Firefox recommended).
2. Either **import** an existing project JSON, or fill in your details manually.
3. Switch to **Generate & Edit** and click **Generate** to build your report.
4. Export via **Print / Save PDF** or **Download .docx**.

---

## Tab-by-Tab Guide

### 1. Setup Tab

This is the starting point. Two modes are available, toggled by the buttons at the top.

#### Import Data Mode

Use this if you have previously exported a project file, or are migrating from an older version of the tool.

| Button | What it accepts | What it restores |
|---|---|---|
| **Choose JSON File** | Full project backup (`.json`) | Profile + Timetable + Holidays |
| **Import Previous Timetable** | Legacy slots-only `.json` | Timetable slots only (merges with existing) |

After a successful full import, the tool automatically switches you to Manual Entry so you can review what was loaded.

> **Note:** The Legacy import is for older exports that only contained timetable slots. It merges the imported slots into your current data rather than replacing everything.

#### Manual Entry Mode

Fill in your personal details and build your weekly timetable from scratch.

**Personal Details fields:**

| Field | Description |
|---|---|
| Teacher Name | Your full name as it should appear on the report |
| Designation | e.g. *Guest Assistant Professor* |
| Department | e.g. *CSE* |
| HOD Name | Head of Department's name (appears in the signature block) |
| HOD Designation | e.g. *Head, Dept of CSE* |

Click **Save Profile** after filling in your details. Profile data is saved to local storage immediately.

**Timetable Slots:**

Add one slot per class session. Each slot requires:

- **Day** — the weekday this class repeats on every week
- **Class / Subject** — the class code or subject name (e.g. `6th Sem B.Tech`, `DAA Lab`)
- **Start / End** — the time window of the class period

Click **Add Slot** to save the entry. It appears in the schedule table below. Use the **×** button to remove any slot. You can add as many slots as needed, including multiple slots on the same day.

---

### 2. Generate & Edit Tab

This is where the report is built and previewed.

#### Report Controls

| Control | Description |
|---|---|
| **Month** | The month to generate the report for |
| **Year** | The year (defaults to current year) |
| **Generate** | Builds the report from your timetable and selected month |
| **Entries per Column** | How many rows appear in each left/right column before a new page starts (default: 20) |
| **Max Total Entries** | Hard cap on total entries across all pages (leave blank for no limit) |
| **Overflow Behaviour** | *Split into multiple tables / pages* — continues on a new page; *Truncate* — stops at the max total |

#### Holidays / Leave Dates

Before generating, add any dates in the selected month that should be excluded from the report (public holidays, leave taken, etc.).

- Pick the date using the date picker
- Enter a label (e.g. `Bihu`, `Medical Leave`)
- Click **Add**

Holidays appear as colour-coded chips. Click **×** on a chip to remove it. These are saved to your project and persist across sessions.

#### Editing the Preview

Once generated, the report renders as a live, editable table in the browser:

- **Class / Subject** and **Time** cells are directly editable — click any cell and type to modify it
- The **+** button on the left of each row inserts a blank entry above that row, shifting all entries below it down
- The **+** button at the bottom of each column appends a new blank row at the end
- Click **Refresh Preview** to re-render the report from the current timetable and layout settings without losing manual edits

The footer of the report shows a working-days summary: total working days, Sundays, Saturdays, and holidays/leave in the selected month.

---

### 3. Project Data Tab

Displays the complete current state of your project as formatted JSON. This includes your profile, all timetable slots, and all holiday entries.

- **Export JSON** — downloads a full project backup file (`RGU_Project.json`) that can be re-imported on any device

> To import data, use the **Setup** tab. The Project Data tab is read-only and for reference/backup only.

---

## Exporting Reports

### Print / Save as PDF

Click **Print / Save PDF** in the Generate & Edit tab. This opens the browser's print dialog with the report pre-configured for **A4 Landscape** orientation. To save as a PDF file instead of printing, select **Save as PDF** (or **Microsoft Print to PDF**) as the destination in the print dialog.

The print stylesheet hides all UI controls and action buttons so only the report content is captured.

### Download as .docx

Click **Download .docx** to generate and download a Word document. The file is named `Duty_Report_<Month>_<Year>.docx`.

The Word document is generated entirely in the browser using the `docx.js` library (loaded from a CDN — an internet connection is required for this feature). It mirrors the layout of the on-screen preview: A4 landscape orientation, Times New Roman font, two-column table, per-column totals, grand total row, and an HOD signature block on the final page.

---

## Data Persistence

All data (profile, timetable, holidays) is stored in the browser's **local storage** under the key `dutyAppUnified`. This means:

- Data survives page refreshes and browser restarts on the same device and browser
- Data is **not** shared between different browsers or devices
- Clearing browser data / site data will erase it — use **Export JSON** regularly to keep a backup

---

## Fixing the .docx Report Layout

If the downloaded `.docx` file does not render in full landscape orientation, or the table columns are too wide / too narrow and overflow the page, apply the following fixes directly in the `exportDocx()` function inside the HTML file.

Open the file in a text editor, locate the `exportDocx` function, and apply the changes described below.

---

### Fix 1 — Correct Landscape Orientation

The `docx.js` library requires portrait dimensions to be passed even when landscape is intended — it swaps width and height internally when `PageOrientation.LANDSCAPE` is set.

Find the `sections.push(...)` block and replace the `properties` object as shown:

```js
// BEFORE (may cause orientation issues in some Word versions)
properties: {
  page: {
    orientation: PageOrientation.LANDSCAPE,
    margin: { top: 567, right: 567, bottom: 567, left: 567 }
  }
}

// AFTER — explicit A4 dimensions with portrait values, landscape flag
properties: {
  page: {
    size: {
      width: 11906,   // A4 short edge in DXA (pass portrait width)
      height: 16838,  // A4 long edge in DXA  (pass portrait height)
      orientation: PageOrientation.LANDSCAPE  // docx.js swaps them in the XML
    },
    margin: { top: 567, right: 567, bottom: 567, left: 567 }
  }
}
```

> **Why this works:** In DXA units (1440 DXA = 1 inch), A4 is 11,906 × 16,838. When `PageOrientation.LANDSCAPE` is set, `docx.js` writes the long edge as the page width in the XML. Passing the values the wrong way round causes some versions of Word to open the document in portrait.

---

### Fix 2 — Fit the Table to a Single Page Width

The table uses six columns (Date, Class/Subject, Time × 2). The current column widths are defined in the `COL` constant and must sum exactly to `CONTENT_WIDTH`.

**Current values (may cause overflow):**
```js
const CONTENT_WIDTH = 13920; // declared content width
const COL = [1740, 4300, 1870, 1740, 4300, 1870]; // sum = 13920 ✓
```

**A4 landscape content width calculation:**

With A4 landscape and 1 cm margins on all sides (567 DXA each):

```
Page width (long edge) = 16838 DXA
Left margin            =   567 DXA
Right margin           =   567 DXA
─────────────────────────────────
Content width          = 15704 DXA
```

Replace both constants with the corrected values:

```js
// AFTER — correct content width for A4 landscape with 1 cm margins
const CONTENT_WIDTH = 15704;

// Distribute across 6 columns so they sum exactly to CONTENT_WIDTH
// Ratio: Date (narrow) | Class/Subject (wide) | Time (medium) — mirrored
// 15704 ÷ 6 ≈ 2617 each; widen the subject columns, narrow the others
const COL = [1800, 5552, 2000, 1800, 5552, 2000]; // sum = 18704 — adjust as needed
```

> **Tip for tuning:** The six values in `COL` must add up **exactly** to `CONTENT_WIDTH`. The "Class / Subject" columns (indices 1 and 4) can be made wider to accommodate long subject names; the "Date" and "Time" columns can stay narrower. Always verify: `1800 + 5552 + 2000 + 1800 + 5552 + 2000 = 18704` — if your sum does not match `CONTENT_WIDTH`, Word will render the table incorrectly.

Also update the `Table` constructor to reference the new width:

```js
// This line already uses CONTENT_WIDTH — no change needed if you updated the constant
children.push(new Table({ width: { size: CONTENT_WIDTH, type: WidthType.DXA }, rows: tableRows }));
```

And update the Grand Total spanning cell, which also references `CONTENT_WIDTH`:

```js
// Already uses CONTENT_WIDTH — automatically correct after updating the constant
new TableCell({
  borders, columnSpan: 6,
  shading: { fill: 'EEEEEE' },
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  ...
})
```

---

### Summary of Changes

| What | Variable / Location | Old value | New value |
|---|---|---|---|
| Page size declaration | `sections.push → properties.page` | `orientation` only | Explicit `size` + `orientation` |
| Content width | `const CONTENT_WIDTH` | `13920` | `15704` |
| Column widths | `const COL` | `[1740, 4300, 1870, 1740, 4300, 1870]` | `[1800, 5552, 2000, 1800, 5552, 2000]` (or your preferred ratio) |

After saving the HTML file, re-open it in the browser and click **Download .docx** again to get the corrected document.

---

## Tips & Troubleshooting

**Report shows 0 entries after generating**
Ensure you have added timetable slots in the Setup tab and that the selected month/year has working days that match those weekdays.

**A date I marked as holiday is still appearing**
Holidays are filtered per month. Make sure the holiday date matches the month and year you are generating for.

**The .docx download does nothing**
The `docx.js` library is loaded from a CDN. Make sure you have an active internet connection when clicking Download .docx. Check the browser console for errors if the button spins indefinitely.

**I edited cells in the preview but the .docx does not reflect them**
Manual cell edits update the in-memory `reportEntries` array, which the `.docx` export reads from. As long as you do not click **Generate** again (which rebuilds the array from scratch), your edits will be preserved in the export.

**My data disappeared after clearing browser history**
Local storage is cleared with browser history. Always export a JSON backup before clearing site data.

**The report splits across too many pages**
Increase the **Entries per Column** value in the report controls. Each page accommodates `Entries per Column × 2` entries (left + right columns). Setting it to 31 with a 31-working-day month will fit all entries on one page.
