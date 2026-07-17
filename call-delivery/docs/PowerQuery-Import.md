# Power Query — Auto-Ingest the Intake Folder

This is the query the **master workbook** uses to pull every group's intake file out of the
shared folder and stack them into one consolidated table. No copy/paste, no per-group tabs.

## One-time setup

1. Open the master workbook → **Data → Get Data → Launch Power Query Editor**.
2. **New Source → Blank Query**, then **Advanced Editor**, and paste the `M` below.
3. Set `FolderPath` to your real shared path (UNC path or a synced OneDrive/SharePoint path).
4. **Close & Load To… → Table** (or **Only Create Connection** + load to the data model).

After that, updating the rollup is just **Data → Refresh All**.

## The query

```m
let
    // 1. Point at the current month's intake folder
    FolderPath = "\\CallDelivery\Intake\2026-07",

    Source     = Folder.Files(FolderPath),

    // 2. Keep real .xlsx files only; drop Excel's temp lock files (~$...)
    OnlyBooks  = Table.SelectRows(Source, each
                    [Extension] = ".xlsx" and not Text.StartsWith([Name], "~$")),

    // 3. From each workbook, grab the table named tblSchedule
    WithData   = Table.AddColumn(OnlyBooks, "Data", each
                    let
                        Book = Excel.Workbook([Content], true),
                        Tbl  = Book{[Item="tblSchedule", Kind="Table"]}[Data]
                    in  Tbl),

    // 4. Expand every group's rows into one long table
    Expanded   = Table.ExpandTableColumn(WithData, "Data",
                    {"Business Group","Date","Start Time (ET)","Timezone Wave",
                     "List ID / Campaign","Priority","Segment / Tag","Special Instructions"}),

    // 5. Drop blank rows (unused template rows have no Date)
    NoBlanks   = Table.SelectRows(Expanded, each [Date] <> null and [Date] <> ""),

    // 6. Type the columns and keep what we need (Name = source filename, handy for audit)
    Renamed    = Table.RenameColumns(NoBlanks, {{"Name", "Source File"}}),
    Keep       = Table.SelectColumns(Renamed,
                    {"Business Group","Date","Start Time (ET)","Timezone Wave",
                     "List ID / Campaign","Priority","Segment / Tag","Special Instructions","Source File"}),
    Typed      = Table.TransformColumnTypes(Keep,
                    {{"Date", type date}, {"Start Time (ET)", type time}, {"Priority", Int64.Type}})
in
    Typed
```

## Notes

- **Named table, not sheet position.** The query looks up `tblSchedule` by name, so it does
  not matter where a group put the table or whether they added rows — as long as they used
  the template and didn't rename the table.
- **New month = one edit.** Change the `2026-07` in `FolderPath` (or parameterize it with a
  workbook cell) and refresh.
- **Bad/renamed files** simply error on that one file; the others still load. Keep an eye on
  the query's error count after a refresh.
- From `Typed`, build the Daily Rollup with a **PivotTable** (rows = Business Group, columns =
  Start Time, values = the wave/campaign) or feed a formatted rollup — that is roadmap step 4.
