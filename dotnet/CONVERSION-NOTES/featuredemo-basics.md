# Conversion notes: featuredemo-basics

Scope: the 12 feature demos CreateWorkbook, OpenWorkbook, SaveExport, ManageWorksheets,
CellReadWrite, BulkRange, FormulasRecalculation, ImportObjects, CellStyling, NumberFormats,
RowsColumns, MergeFreeze, plus `FeatureDemoBasicsIntegrationTest`.

## How the port was checked

- Ran all 12 Java demos (`./mvnw -q -o exec:java -Dexec.args="feature-demo/<name> <dir>"`) and
  all 12 .NET demos (`dotnet run --project src/Axyra.Sheets.Examples -- feature-demo/<name> <dir>`),
  then compared every output file. For each `.xlsx`, every zip part except `docProps/*` was diffed
  after removing the random `xr:uid` GUIDs.
- **Verified:** the console output matches line for line. The one exception is the `.xlsx` byte
  count in save-export: Java 4652, .NET 4651. Everything except `docProps` and the `xr:uid` GUIDs
  is identical, so the byte must come from `docProps` (timestamps or app metadata). I did not
  check which field.
- **Verified:** `quarterly-report.csv`, `.json` and `.html` are byte-identical across the two SDKs.
- **Verified:** every workbook has the same structure, values, formulas, cached results, styles,
  merges, freeze panes, tab colours, widths and heights in both SDKs. 17 of the 45 workbooks are
  identical apart from those parts. In the other 28, the only differences are in
  `sharedStrings.xml`, and each one is a label or a stringified value listed under D1 and D2 below.
- `FeatureDemoBasicsIntegrationTest` passes (`./run-tests.sh FeatureDemoBasicsIntegrationTest`: 1/1).
  No test needed a known-issue mark.

No .NET SDK bug turned up in this area. The findings below are two problems in the Java examples
(the .NET port reproduces them), plus design differences recorded for completeness.

## F1. Java example bug (reproduced faithfully): ImportObjects labels do not match the columns

- Java: `ImportObjectsExample.Trip` is a JavaBean, and the Java SDK orders bean getters **by property
  name** (`DataImport.java`: "JavaBean getters ... sorted by property name"). The columns therefore
  come out as `city, cost, nights`, but `beans()` and `importAndTotal()` pass
  `columnLabels("City", "Nights", "Cost")`. The header "Nights" sits over the cost column and
  "Cost" over the nights column. In `importAndTotal()`, "Cost per night" = `C7/B7` = nights/cost.
- .NET: `DataImport.Discover` orders a non-record class's public properties by name too (ordinal
  sort). The faithful port (a plain class with properties `City`, `Nights`, `Cost`) gives the same
  output.
- Observed (identical in both SDKs, `import-objects-total.xlsx`, sheet Trips):
  `A1 City | B1 Nights | C1 Cost | A2 Lisbon | B2 820 | C2 4 | ... | B7 =SUM(B2:B5) -> 4660 | C7 =SUM(C2:C5) -> 16 | C8 =C7/B7 -> 0.00343...`
- Classification: bug in the Java example, not in either SDK. Fix: either `columns("city", "nights", "cost")`
  (.NET: `Columns = ["City", "Nights", "Cost"]`) or labels in alphabetical order. I left it
  unfixed so the port stays faithful.
- Related, verified from the source: a C# `record Trip(string City, int Nights, double Cost)`
  imports in **declaration** order (`IsRecord` -> `OrderBy(MetadataToken)`), as a Java record does
  (component order). So the Java comment "A record works identically" does not hold in either
  language: with a record, the labels would line up.

## F2. Java example and doc inconsistency: "a formula is inert until recalculation"

- Java `FormulasRecalculationExample` says "Setting a formula does not evaluate it —
  `recalculate()` does", and `beforeAndAfter()` is meant to show that D2 has no result before
  `recalculate()`.
- Observed in **both** SDKs (`formulas-before-after.xlsx`): "D2 before recalculate()" already
  reads `Number[value=59.97]` (Java) / `Number { Value = 59.97 }` (.NET). Setting the formula
  evaluates it immediately. The .NET QuickStart says the same thing ("calculated right away").
- Classification: doc gap in the Java example's comment, which behaviour contradicts. The engine
  behaves the same through both SDKs. I kept the faithful code and the original comment. The
  sheet output shows the real behaviour.

## D1. Labels that name a Java API were renamed to the .NET API (intentional, port-level)

Several demos write the name of the call into a cell, for example `fontName("Georgia")`,
`insertSheet(0, "Cover")`, `setNumbers(double[])`, `autoFitColumn(0)`, `isRowHidden(3)`,
`Formats.format(...)`, `Workbook.createColumnar()`, "single JNI crossing" and "computed in Java".
The port writes the .NET member instead (`Font.Name = "Georgia"`, `Sheets.Insert(0, "Cover")`,
`SetNumbers(ReadOnlySpan<double>)`, `Columns[0].AutoFit()`, `Rows[3].Hidden`, `Formats.Format(...)`,
`new Workbook(WorkbookStorage.Columnar)`, "single native crossing" and "computed in C#").
CellReadWrite's row label `setNumber` became `implicit`, because .NET has no `Sheet.setNumber`
(see D3). Every changed string is in the diff summary above. Nothing else differs.

## D2. Stringified values differ by language (intentional design and language difference)

| Where | Java | .NET | Why |
|---|---|---|---|
| `String.valueOf(CellValue)` (formulas ad-hoc, before/after) | `Number[value=7.0]`, `Text[value=12.3%]` | `Number { Value = 7 }`, `Text { Value = 12.3% }` | Java record `toString` vs C# record `ToString` |
| `wb.calcMode()` | `MANUAL` | `Manual` | .NET enums use PascalCase (16-design-conventions) |
| `ParsedValue.kind()` | `NUMBER`, `TEXT` | `Number`, `Text` | same |
| doubles (`fontSize`, `columnWidth`, `rowHeight`, `defaultRowHeight`, `String.valueOf(1.0)`) | `11.0`, `26.0`, `36.0` | `11`, `26`, `36` | `double.ToString()` drops `.0` |
| booleans (`isRowHeightCustom`, `isRowHidden`) | `true` | `True` | `bool.ToString()` |
| `LocalDateTime.toString()` vs `DateTime.ToString("s")` | `2026-08-24T14:30` | `2026-08-24T14:30:00` | the port uses ISO `"s"`. `DateOnly` with `"O"` gives `2026-08-24`, the same as Java |

The values underneath are identical: same widths, same serials, same cached results.
I formatted the doubles and dates with `CultureInfo.InvariantCulture`, so a non-English machine
does not change them.

## D3. API mapping used (no capability missing for these demos)

| Java | .NET |
|---|---|
| `Workbook.create()` / `createColumnar()` | `new Workbook()` / `new Workbook(WorkbookStorage.Columnar)` |
| `Workbook.open(InputStream, "xlsx")` / `openBytes(byte[], "xlsx")` | `Workbook.Open(Stream, WorkbookFormat.Xlsx)` / `Workbook.Open(byte[], WorkbookFormat.Xlsx)` |
| `wb.saveBytes("csv")` | `wb.ToByteArray(WorkbookFormat.Csv)`; file name from `format.Extension()` |
| `createSheet` / `insertSheet` / `renameSheet(i, n)` / `moveSheet` / `copySheet` / `deleteSheet(i)` | `Sheets.Add` / `Sheets.Insert` / `sheet.Name = n` / `Sheets.Move` / `Sheets.AddCopy` / `Sheets.RemoveAt` |
| `wb.setActiveSheet(int)` | `wb.ActiveSheet = sheet` (takes a `Sheet`, not an index) |
| `wb.setDefaultFont(CellStyle)` / `defaultStyle().fontName()` | `wb.DefaultFont = new Font {...}` / `wb.DefaultStyle.Font.Name` |
| `sheet.setNumber(r, c, d)` | `sheet[r, c].Value = d` (implicit conversion; there is no `SetNumber`) |
| `range.setNumbers(double[])` (flat, row-major) / `range.numbers()` (flat) | `SetNumbers(ReadOnlySpan<double>)` (flat, row-major) / `GetNumbers()` returns **`double[,]`** (2-D), so the totals loop walks rows |
| `wb.evaluate(expr, sheet)` | `sheet.Evaluate(expr)` |
| `Functions.count()` | `BuiltInFunctions.Count` |
| `setColumnWidth` / `setRowHeight` / `setDefaultRowHeight` / `autoFitColumn` / `insertRows` | `Columns[i].Width` / `Rows[i].Height` / `Rows.DefaultHeight` / `Columns[i].AutoFit()` / `Rows.Insert` |
| `CellStyle.builder().allBorders(s, c)` | `Borders = Borders.Around(s, c)` (verified in the Java source: both set the four edges of every cell) |
| `ImportOptions.builder().columns("cost", "city")` | `Columns = ["Cost", "City"]` (see D4) |

## D4. `ImportOptions.Columns` matches .NET property names exactly (intentional)

- Java: bean property names are camelCase (`"cost"`, `"city"`).
- .NET: discovery keys are the C# property names, compared case-sensitively (`StringComparer.Ordinal`).
- Repro:
  ```csharp
  s.ImportData(0, 0, trips, new ImportOptions { Columns = ["cost", "city"] });
  ```
- Observed: `AxyraException: unknown column 'cost'; available: [City, Cost, Nights]`.
- Classification: intentional design difference, because names follow each language's
  conventions. The error message lists the valid names. The port uses `["Cost", "City"]`.
