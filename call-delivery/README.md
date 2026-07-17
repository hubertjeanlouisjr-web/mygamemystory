# Call Delivery — Streamlined Intake

This folder holds the standardized **outbound dialing schedule** intake process for the
Call Delivery support team. It replaces the manual "groups email us / we retype into Excel"
workflow with a single standard template + a shared folder that a master workbook imports
automatically via Power Query.

## Contents

| Path | What it is |
|---|---|
| `templates/Call_Delivery_Intake_Template.xlsx` | The standard template every business group fills in. Locked format, dropdown-driven, one row per scheduled pass. |
| `intake/` | The central drop location. One subfolder per GROUP, files named by month (`FE/2026-08.xlsx`). |
| `docs/PowerQuery-Import.md` | The Power Query (M) code the master workbook uses to ingest every group folder under `intake/` in one refresh. |

## How the process works

```
Business groups                Central shared folder                 Call Delivery master workbook
─────────────────              ─────────────────────                 ─────────────────────────────
Fill the standard    ──save──▶  \CallDelivery\Intake\      ──Power──▶  Refresh → all groups appended
template (dropdowns)            FE\2026-08.xlsx             Query       → Daily Rollup + Assignments
                                BE\2026-08.xlsx                          (no copy/paste, no #REF!)
                                Card\2026-08.xlsx  ...
```

## Central folder layout & naming

```
\CallDelivery\Intake\
    FE\
        2026-07.xlsx
        2026-08.xlsx
    BE\
        2026-08.xlsx
    Card\
        ...one folder per business group...
```

**Rules**

- One folder **per group**; grant each group write access to **their folder only** —
  nobody can overwrite another group's schedule.
- Files are named by month (`YYYY-MM.xlsx`). When a group's schedule changes, they
  **overwrite the same file** — never add a "v2" copy (two files for one month both
  import; the Source File column exposes it).
- The master's query points at the Intake root once and filters by the month cell on
  its Settings tab — no path edits at month roll-over.
- Groups never touch the master workbook. They only ever open their own template file.

## Why the template is shaped the way it is

The template stores the schedule as a clean **Excel Table (`tblSchedule`)** in a *long*
format — one row per pass, with real Date and Time values and dropdown-controlled codes.
That is exactly what Power Query needs to combine 15–20 files reliably. The old workbook
stored schedules "wide" (an hourly column-pair grid) with times as text like `"8:15a"`,
which is why it depended on a fragile web of cross-sheet formulas (219 `#REF!` errors in
the Daily Rollup, plus an external-workbook link and a fully-dead `Auto` row).

## Roadmap (agreed direction)

1. **Standard intake template** ✅ *(this deliverable)* — Excel + Power Query.
2. **Master workbook** — Power Query ingest of `intake/<month>/` → consolidated table.
3. **Assignment randomizer** — fair-rotation staffing generator (balances load, avoids the
   same person repeating on the same group).
4. **Redesigned Daily Rollup** — same information, clean card/timeline layout, no error cells.
