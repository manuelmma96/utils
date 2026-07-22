# buffer-report-tool

Standalone tool to analyze GPS packet JSON files and generate a consolidated
Excel report of "buffer" gaps between the moment a device took a GPS fix and
the moment that data actually arrived/was saved to the database.

It doesn't depend on any backend service or a Mongo connection — it reads
exported JSON files directly (e.g. a collection dump).

## Input data requirements

Each document in the JSON must have, at minimum:

- A Mongo ObjectId in `_id.$oid` — the **End** date is derived from it
  (see "How the buffer is calculated" below).
- One or more records with a `timestamp` field inside
  `data.dataRecords[]` — the **Start** date (the GPS fix) comes from there.

If your JSON has those two elements under those field names, the script
works unmodified regardless of which device/IMEI it came from. Just drop it
in as another `.json` file inside the input folder (`GPS_PACKETS/` by
default) and run it — nothing is hardcoded to a specific IMEI or file.

## How the buffer is calculated

MongoDB's ObjectId embeds a creation timestamp in its first 4 bytes (8 hex
characters). Since the documents in this dataset have no explicit
`createdAt` field, it's derived from there:

```js
const seconds = parseInt(oid.substring(0, 8), 16);
const end = new Date(seconds * 1000);
```

- **Start** (`Inicio`) = `data.dataRecords[i].timestamp` (as-is from the device)
- **End** (`Fin`) = date derived from `_id` (see above)
- **Duration (minutes)** (`Duracion (minutos)`) = `(End - Start) / 60000`

Only records whose duration exceeds the configured threshold (5 minutes by
default) are listed in the detail sheet. If the duration is negative (the
fix is later than the derived date), it's counted separately as an anomaly —
that usually indicates clock skew, not a real buffer.

## Installation

```bash
cd buffer-report-tool
npm install
```

## Usage

```bash
node index.js [options]
```

| Flag | Default | Description |
|---|---|---|
| `--input <dir>` | `../GPS_PACKETS` | Folder with the `.json` files to process |
| `--output <file>` | `../output/reporte_consolidado_gps.xlsx` | Output Excel path (folder created automatically if missing) |
| `--threshold <min>` | `5` | Gap threshold in minutes to count as "buffer" |
| `--from <ISO date>` | (none) | Only consider records with Start >= this date |
| `--to <ISO date>` | (none) | Only consider records with Start < this date |

Without `--from`/`--to`, no date filter is applied — every record in the
file is processed. To restrict to a specific month:

```bash
node index.js --from 2026-07-01T00:00:00.000Z --to 2026-08-01T00:00:00.000Z
```

`npm start` runs the same thing as `node index.js`; to pass flags through
`npm start` you need the `--` separator:

```bash
npm run start -- --threshold 10
```

## What the Excel report contains

- **Resumen Ejecutivo** (Executive Summary): one row per processed file
  (total documents, total records, buffered count, % buffered, max/average/
  median gap, anomalies, duration bucket distribution) plus a TOTAL row.
- **Buffer_`<imei>`**: one sheet per file with the detail of every record
  that exceeded the threshold — columns `IMEI`, `doc_id`, `record_index`,
  `Inicio` (Start), `Fin` (End), `Duracion (minutos)` (Duration in minutes),
  `latitude`, `longitude`, `speed`, `ignition`. If a file produces more rows
  than fit in a single Excel sheet, it's automatically split into `_pt2`,
  `_pt3`, etc.
- Every column header carries a tooltip (Excel comment) explaining where
  that value comes from / how it's calculated.
- The report's sheet names, headers, and tooltips are written in Spanish,
  since that's the language of the intended audience for this report. Dates
  are written as plain ISO 8601 text (not an Excel "date" cell), so they
  display identically regardless of the viewer's regional settings and
  don't require clicking to expand.

## Memory safety with large files

Source JSON files can be tens of MB with millions of lines. The script
never runs `JSON.parse` on the full file: it uses `JSONStream` to parse
document by document in streaming mode, and only keeps in memory the rows
that exceed the buffer threshold (a small fraction of the total). The Excel
file is also written in streaming mode (`exceljs`'s `WorkbookWriter`).
Processing a ~90MB / 3M-line file takes ~1-2 seconds with a peak RAM usage
of ~90-100MB.

## Malformed JSON detection

Before processing each file, the script inspects its last bytes:

- If the file correctly ends in `]`, it proceeds without any warning.
- If the closing `]` is missing but the last document closes with `}` (a
  complete document, only the root array's closing bracket is missing), a
  console warning is printed and **the file is processed anyway** — no
  document is lost because of this.
- If the cut happened mid-structure (ending in neither `}` nor `]`), the
  report generation is **blocked** with a clear error, since in that case a
  document could genuinely have been lost mid-write without any parse error.

## Cross-validation

This script's results were verified by independently running the same
calculation with `jq` (C) and with Python (native `json`, no streaming) on
the same files, getting identical counts and buffer durations across all
three. See this folder/conversation's history for the details of that
validation if you need to repeat it.
