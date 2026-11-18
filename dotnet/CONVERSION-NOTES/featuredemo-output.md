# Conversion notes: feature demos, data and output

Area: `FindReplace`, `SortFilter`, `DataValidation`, `ConditionalFormat`, `CreateTable`,
`InsertChart`, `ExportPdf`, `RenderImage` (`src/Axyra.Sheets.Examples/FeatureDemo/*Example.cs`)
and `tests/Axyra.Sheets.Examples.Tests/FeatureDemoDataAndOutputIntegrationTest.cs`.

Package: `Axyra.Sheets 1.0.0-fl.20261116` (.NET 8). Java side: `axyra-sheets-examples`,
run with `./mvnw -q exec:java -Dexec.args="feature-demo/<name> <dir>"`.

## How the outputs were compared

I ran every demo through both CLIs into separate directories and compared the outputs part by part.
- **Console output:** the same, line for line, in all 8 demos (including the byte counts
  printed for the PDFs and images).
- **PDF, PNG and SVG:** the same bytes as Java's (`report.pdf`, `report-sheet1.pdf`,
  `report-fit-width.pdf`, `sheet.png`, `sheet-2x.png`, `sheet-no-gridlines.png`, `sheet.svg`).
  - `pdfinfo`: `report.pdf` has 2 A4 pages (Summary, By rep), `report-sheet1.pdf` has 1 and
    `report-fit-width.pdf` has 2.
  - PNG sizes are 1123x794 (scale 1, landscape A4 at 96 DPI), 2245x1587 (x2) and 1684x1191 (x1.5).
  - I looked at the images and rendered PDF pages. They match what the demos intend: the table
    and the column chart sit side by side on one landscape page, and there are no gridlines
    when `Gridlines = false`. The PDF has the title band, the table, the chart below it, and a
    "By rep" page.
- **XLSX:** I compared every part after removing the random `xr:uid` of the evaluation sheet.
  All 36 workbooks are the same as Java's except for the three differences below
  (D1, D2, D3). The computed values (hit counts, rows hidden, table totals, page counts) are
  the same.

## D1. Data validation: .NET writes `showInputMessage="1"` on every rule (design difference)

- **Java:** `Validation.wholeNumber("between", "1", "100")` and the other factories write
  `<dataValidation ... allowBlank="1" showErrorMessage="1">`. `showInputMessage` is written
  only when an input message is set (the `prompts` workbook).
- **.NET:** `Validation.WholeNumber(ValidationOperator.Between, "1", "100")` writes
  `allowBlank="1" showInputMessage="1" showErrorMessage="1"`.
- **Repro:**
  ```csharp
  using Workbook wb = new Workbook();
  Sheet s = wb.Sheets.Add("S");
  s.Validations.Add("A1:A3", Validation.WholeNumber(ValidationOperator.Between, "1", "100"));
  Console.WriteLine(s.Validations[0].Rule.ShowInputMessage);   // True
  ```
- **Observed:** this affects `data-validation-dropdown`, `-numeric` and `-custom.xlsx` (one
  attribute on each rule). `-prompts.xlsx` is the same in both.
- **Classification:** intentional design difference, documented. `docs/dotnet-api/09` says
  "`AllowBlank`, `ShowErrorMessage` and `ShowInputMessage` are on by default". With no
  title or message set, Excel shows nothing either way, so users see no difference. The
  default differs between the SDKs, though, and a parity check that compares bytes would
  flag it.

## D2. Chart legend "none": Java writes a deleted legend, .NET omits it (missing API / design difference)

- **Java:** `.legend(LegendPosition.NONE)` writes `<c:legend><c:legendPos val="b"/>…<c:delete val="1"/></c:legend>`.
  Java also has `legendVisible(false)`, which leaves the element out.
- **.NET:** `LegendPosition` has no `None` (it has `Right`, `Top`, `Bottom`, `Left` and
  `TopRight`). The only way to hide the legend is `chart.Legend.Visible = false`. That
  leaves `<c:legend>` out completely, like Java's `legendVisible(false)`.
- **Port:** `ExportPdfExample.Report()` uses `chart.Legend.Visible = false`.
- **Observed:** `xl/charts/chart1.xml` of `export-pdf-report.xlsx` and
  `export-pdf-page-count.xlsx` is missing the deleted-legend block. The rendered PDF is the
  same, with no legend in either.
- **Classification:** small missing API (.NET cannot write the "deleted legend" form). It is
  harmless, because both forms hide the legend in Excel and in the renderer.

## D3. Data bar axis: the Java demo writes the invalid token `midpoint`; .NET writes `middle` (Java-side bug)

- **Java demo:** `.axisPosition("midpoint")`. The Java builder takes a free string. Its
  Javadoc says `"automatic"`, `"middle"`, or `"none"`, but it does not validate the value, so
  the file gets `<x14:dataBar axisPosition="midpoint">`. That is not a value of
  `ST_DataBarAxisPosition` (`automatic | middle | none`). The engine's own test
  `crates/axyra-sheets-ooxml/tests/bor_ft_025.rs` uses `middle`.
- **.NET:** `AxisPosition = DataBarAxisPosition.Middle` writes `axisPosition="middle"`. The
  enum makes the Java mistake impossible.
- **Observed:** this is the only difference in `conditional-format-data-bars.xlsx`.
- **Classification:** a bug in the Java example, plus a missing validation in the Java SDK.
  It is not a .NET issue, and the .NET output is the correct one. I have not checked what
  Excel does with `midpoint`.

## D4. Find/replace and imported cells: the Java demo's reason for rewriting is out of date (doc gap, Java example)

- The Java `tickets()` reads the used block and writes it straight back, "because find and
  replaceAll match only cells that have been written through the Java API".
- **Verified, both SDKs, without the rewrite:** `find("overdue")` gives 6,
  `findRegex("^INV-\d{4}$")` gives 8 and `replaceAll("overdue", "X")` gives 6, the same as
  with the rewrite. I ran a Java probe (`Workbook.openBytes` + `sheet.find`) and a .NET probe
  (`Workbook.Open(bytes, Xlsx)` + `sheet.FindAll`) against `findReplace-tickets-content.xlsx`.
- **Port:** the rewrite is kept so the steps stay the same as Java's. The C# comment says it
  is no longer needed.
- **Classification:** doc gap in the Java example's comment. This is not a .NET difference.

## API mapping notes (no behaviour difference)

| Java | .NET |
|---|---|
| `sheet.find(t)` → `int[][]` | `sheet.FindAll(t)` → `IReadOnlyList<Cell>` |
| `find(t, case, entire, wildcard)` / `findRegex` / `replaceAllRegex` | `FindOptions { MatchCase, MatchEntireCell, UseWildcards, UseRegex }` |
| `setAutoFilter(r)`, `setColumnFilter(0,"East")`, `setColumnCustomFilter`, `applyAutoFilter()` | `AutoFilter.Range = r`, `AutoFilter[0].FilterByValues("East")`, `FilterByCondition`, `AutoFilter.Apply()` (the criteria already take effect when set; `Apply()` returns the same hidden count Java prints: 18 and 13) |
| `range.createTable(name, cols…)`, `setTableHeaderRow/TotalsRow(name, b)` | `sheet.Tables.Add(name, range, cols…)`, `Tables[name].ShowHeaderRow/ShowTotalsRow` |
| `Chart.column()…anchor(r1,c1,r2,c2)` | `sheet.Charts.Add(ChartType.Column, new DrawingAnchor(r1,c1,r2,c2))`. Java's anchor is a pair of *corner* cells, so it maps to `new DrawingAnchor(...)`, **not** to `DrawingAnchor.Cover(...)` or an A1 range, which cover the cells inclusively and would make every chart one row and one column larger |
| `.series(name, v, c, color)` | `Series.Add(name, v, c).FillColor = color` |
| `wb.sheetPageCount(i)`, `Renderer.pageCount(wb, i)` | both `wb.Sheets[i].GetPageCount()` (.NET has one member, so the demo's third row calls it again) |
| `Renderer.toPdf(wb, PdfOptions.builder().sheets(0).fitTo(WIDTH))` | `wb.ToPdf(new PdfOptions { Sheets = [0], FitTo = PdfFit.Width })` |
| `Renderer.toImage(wb, 0, "png", opts)` / `toSvg` | `wb.Sheets[0].RenderImage(ImageFormat.Png / Svg, new ImageOptions { Scale, Gridlines })` |
| `CellStyle.builder().allBorders(THIN, c)` | `Borders = Borders.Around(BorderStyle.Thin, c)` (it is applied to every cell of a range, so each cell gets all four edges, as in Java) |
| `setLandscape(true)`, `fitToPage(1,1)`, `setPrintArea` | `PageSetup.Orientation = Landscape`, `PageSetup.FitToPages(1,1)`, `PageSetup.PrintArea` |

Test port: the Java `LinkedHashMap` is a `Dictionary<string, string[]>`. It is never
removed from, so it keeps insertion order. `OrderedDictionary<,>` would need .NET 9.
