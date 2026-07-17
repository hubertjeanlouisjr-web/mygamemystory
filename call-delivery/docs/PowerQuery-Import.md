# Power Query — Auto-Ingest the Intake Folders

This is the query the **master workbook** uses to pull every group's schedule out of the
shared Intake folders and stack them into one consolidated table. It reads **every group's
subfolder automatically** and filters to the month named on the master's Settings tab —
so the query is set up once and **never edited again**, not even at month roll-over.

## Folder layout (group-first)

```
\\CallDelivery\Intake\
    FE\2026-07.xlsx        2026-08.xlsx ...
    BE\2026-07.xlsx        ...
    Card\...
    ...one folder per business group...
```

**Rules**

- One folder per group; grant each group write access to **their folder only**.
- Files are named by month: `YYYY-MM.xlsx`. When the schedule changes, the group
  **overwrites the same file** — never adds a "v2" copy.
- If a folder ever holds two files for the same month (`2026-08.xlsx` and
  `2026-08 v2.xlsx`), **both** import — the `Source File` column in the output makes
  this immediately visible. Delete the extra file.

## One-time setup

1. Open the master workbook → **Data → Get Data → Launch Power Query Editor**.
2. **New Source → Blank Query**, then **Advanced Editor**, and paste the `M` below.
3. Set `IntakeRoot` to your real shared path (UNC path or synced OneDrive/SharePoint path).
4. **Close & Load To… → Table**, replacing the sample table on the Consolidated tab.

After that, month roll-over is: change the **Month to load** cell on the Settings tab
(the `IntakeMonth` named range) → **Data → Refresh All**. That's it.

## The query

```m
let
    // 1. The Intake ROOT - set once, never changes month to month
    IntakeRoot = "\\CallDelivery\Intake",

    // 2. The month to load (e.g. "2026-08") - read from the master's
    //    Settings tab via the IntakeMonth named range
    Month      = Excel.CurrentWorkbook(){[Name="IntakeMonth"]}[Content]{0}[Column1],

    // 3. Read every group subfolder under the root
    Source     = Folder.Files(IntakeRoot),

    // 4. Keep this month's real .xlsx files; drop Excel temp lock files (~$...)
    OnlyBooks  = Table.SelectRows(Source, each
                    [Extension] = ".xlsx"
                    and not Text.StartsWith([Name], "~$")
                    and Text.StartsWith([Name], Month)),

    // 5. From each workbook, grab the table named tblSchedule
    WithData   = Table.AddColumn(OnlyBooks, "Data", each
                    Excel.Workbook([Content], true){[Item="tblSchedule", Kind="Table"]}[Data]),

    // 6. Expand every group's rows into one long table
    Expanded   = Table.ExpandTableColumn(WithData, "Data",
                    {"Business Group","Date","Start Time (ET)","Timezone Wave",
                     "List ID / Campaign","Priority","Segment / Tag","Special Instructions"}),

    // 7. Drop blank rows (unused template rows have no Date)
    NoBlanks   = Table.SelectRows(Expanded, each [Date] <> null and [Date] <> ""),

    // 8. Source File = "GroupFolder\file.xlsx" so ownership and duplicate
    //    files are visible at a glance
    WithSrc    = Table.AddColumn(NoBlanks, "Source File", each
                    Text.AfterDelimiter([Folder Path], IntakeRoot & "\") & [Name]),

    Keep       = Table.SelectColumns(WithSrc,
                    {"Business Group","Date","Start Time (ET)","Timezone Wave",
                     "List ID / Campaign","Priority","Segment / Tag",
                     "Special Instructions","Source File"}),
    Typed      = Table.TransformColumnTypes(Keep,
                    {{"Date", type date}, {"Start Time (ET)", type time},
                     {"Priority", Int64.Type}})
in
    Typed
```

## Notes

- **Named table, not sheet position.** The query finds `tblSchedule` by name, so groups
  can add rows freely — as long as they use the template and don't rename the table.
- **Per-group permissions are the point** of the group-first layout: nobody can overwrite
  another group's schedule, accidentally or otherwise.
- **A bad or renamed file** errors on that one file only; every other group still loads.
- From `Typed`, the Daily Rollup tab in the master reads the consolidated table
  automatically — pick a Strategy Date and the morning sheet assembles itself.
