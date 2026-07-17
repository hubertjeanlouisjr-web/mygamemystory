# Call Delivery — Streamlined Intake

This folder holds the standardized **outbound dialing schedule** intake process for the
Call Delivery support team. It replaces the manual "groups email us / we retype into Excel"
workflow with a single standard template + a shared folder that a master workbook imports
automatically via Power Query.

## Contents

| Path | What it is |
|---|---|
| `templates/Call_Delivery_Intake_Template.xlsx` | The standard template every business group fills in. Locked format, dropdown-driven, one row per scheduled pass. |
| `intake/` | The central drop location. One subfolder per month, one file per group. |
| `docs/PowerQuery-Import.md` | The Power Query (M) code the master workbook uses to ingest the whole `intake/<month>/` folder in one refresh. |

## How the process works

```
Business groups                Central shared folder                 Call Delivery master workbook
─────────────────              ─────────────────────                 ─────────────────────────────
Fill the standard    ──save──▶  \CallDelivery\Intake\      ──Power──▶  Refresh → all groups appended
template (dropdowns)            2026-07\FE.xlsx             Query       → Daily Rollup + Assignments
                                2026-07\BE.xlsx                          (no copy/paste, no #REF!)
                                2026-07\PCO.xlsx  ...
```

## Central folder layout & naming

```
\CallDelivery\Intake\
    2026-07\
        FE.xlsx
        BE.xlsx
        PCO.xlsx
        Repo.xlsx
        ...one file per business group...
    2026-08\
        ...
```

**Rules**

- One file **per group per month**. Filename = the group's short name (matches the
  `Business Group` dropdown), e.g. `FE.xlsx`.
- When a group's schedule changes, they **overwrite the same file** — the next master
  refresh picks it up.
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
