# regression-features: FeatureCoreTest, FeatureDataTest, FeatureContentTest, FeatureLayoutTest, DocumentedBehaviourTest

Java sources: `axyra-sheets-examples/java/src/test/java/io/keikai/axyra/examples/regression/`
`FeatureCoreTest.java`, `FeatureDataTest.java`, `FeatureContentTest.java`, `FeatureLayoutTest.java`,
`DocumentedBehaviourTest.java`.
.NET ports: `tests/Axyra.Sheets.Examples.Tests/Regression/<same name>.cs`. One `[Fact]` per Java `@Test`,
same method name in PascalCase; the Java `@DisplayName` is kept as the method's `<summary>`.

Package: `Axyra.Sheets 1.0.0-fl.20261116` (evaluation mode).

## Result

| Class | Run | Passed | Failed | Skipped | Known issue |
|---|---|---|---|---|---|
| FeatureCoreTest | `./run-tests.sh FeatureCoreTest` | 13 | 0 | 0 | 0 |
| FeatureLayoutTest | `./run-tests.sh FeatureLayoutTest` | 10 | 0 | 0 | 0 |
| FeatureContentTest | `./run-tests.sh FeatureContentTest` | 12 | 0 | 0 | 0 |
| FeatureDataTest | `./run-tests.sh FeatureDataTest` | 15 | 0 | 1 (`DigitalSignature`, needs a licence, as in Java) | 1 (`FormulaNotation`) |
| DocumentedBehaviourTest | `./run-tests.sh DocumentedBehaviourTest` | 16 | 0 | 0 | 0 |

`./run-tests.sh --known-issues FeatureDataTest` runs `FormulaNotation` and it fails as described in issue 1.
Each class stays well under the 100 opens/saves/renders per process limit.

`IssueTestSupport.cs` was used unchanged; no bug found in it.

## Known issues

### 1. `FormulaToR1C1` / `FormulaToA1` anchor is 1-based in the package, documented as zero-based

- **Java**: `wb.formulaToR1C1("A1+$B$2+SUM(A1:A3)", 3, 3)` (the test calls the anchor 1-based) gives
  `R[-2]C[-2]+R2C2+SUM(R[-2]C[-2]:RC[-2])`, relative to C3.
- **.NET documentation**: `docs/dotnet-api/12-formulas-and-calculation.md`, "A1 ⇄ R1C1 conversion": the
  anchor is "given by its zero-based row and column as `sheet[row, column]` counts them";
  example `wb.FormulaToR1C1("A1+$B$2", 2, 2) // "R[-2]C[-2]+R2C2"` (anchor C3) and
  `wb.FormulaToR1C1("SUM(A1:B2)", 0, 0) // "SUM(RC:R[1]C[1])"`. The SDK source
  (`sdk-dotnet/src/Axyra.Sheets/Workbook.Document.cs`, `ConvertFormula`) says the same and adds 1 before
  calling the engine ("The engine counts the anchor from 1").
- **.NET package behaviour**: the shipped package treats the anchor as 1-based. Its XML doc comment says
  "relative to the 1-based (anchorRow, anchorColumn)", so the package and the current docs/source disagree.
- **Repro** (console app referencing the package):
  ```csharp
  using var wb = new Workbook();
  Console.WriteLine(wb.FormulaToR1C1("A1+$B$2", 2, 2));
  Console.WriteLine(wb.FormulaToA1("R[-2]C[-2]+R2C2", 2, 2));
  Console.WriteLine(wb.FormulaToR1C1("SUM(A1:B2)", 0, 0));
  Console.WriteLine(wb.FormulaToR1C1("A1+$B$2+SUM(A1:A3)", 3, 3));
  ```
  Observed:
  ```
  R[-1]C[-1]+R2C2                          (documented: R[-2]C[-2]+R2C2)
  R[-2]C[-2]+$B$2                          (documented: A1+$B$2; the relative refs fall off the grid and stay R1C1 text)
  SUM(R[1]C[1]:R[2]C[2])                   (documented: SUM(RC:R[1]C[1]); no exception for the off-grid anchor)
  R[-2]C[-2]+R2C2+SUM(R[-2]C[-2]:RC[-2])   (1-based anchor C3 = the documented zero-based (2, 2) result)
  ```
- **Classification**: SDK bug in the shipped package (verified). The newer SDK source already converts
  the anchor to 1-based, so this is probably fixed in a later build (not verified; only the package was run).
  Anyone writing code from the .NET docs against this package gets references shifted by one row and column.
- **Test**: `FeatureDataTest.FormulaNotation` uses the documented zero-based anchor `(2, 2)` and carries
  `[Trait("Category", "known-issue")]`. The rest of the test (the German localized formula) was checked
  separately and works: `SetLocalizedFormula("SUMME(A1:A3)", "de-DE")` stores `SUM(A1:A3)` and gives 60.

## Skipped tests

- `FeatureDataTest.DigitalSignature`: as in Java (`assumeTrue(License.status().isLicensed())`) it is skipped
  through `Assert.SkipUnless(License.Status.IsLicensed, ...)` on the evaluation build. Java runs `keytool` to
  make a PKCS#12 key; the .NET port makes the same self-signed RSA-2048/SHA-256 certificate in process with
  `CertificateRequest` and passes `rsa.ExportPkcs8PrivateKey()` and `cert.RawData` (PKCS#8 DER and X.509 DER,
  what `Workbook.Sign` accepts). Not run, so unverified.

No test needed a `Skip="no .NET counterpart"`: every Java API used here has a .NET equivalent.

## API mapping that changed the shape of a test (not defects)

All verified by compiling against the package and running the tests.

| Java | .NET | Note |
|---|---|---|
| `Workbook.create()`, `createSheet`, `sheet(name)` | `new Workbook()`, `wb.Sheets.Add`, `wb.Sheets[name]` | |
| `cell.setDate(LocalDate)` / `setDateTime(LocalDateTime)` | `cell.DateValue = DateOnly` / `cell.DateTimeValue = DateTime` | |
| `wb.evaluate(f, sheet)` | `sheet.Evaluate(f)` | |
| `wb.addName` / `names()` / `removeName` / `findName == null` | `wb.Names.Add` / enumerate `wb.Names` / `Names.Remove` / `Names.TryGet == false` | |
| `CellStyle.builder()...build()`, `style.toBuilder()` | object initializer, `style with { ... }` | `fillColor()` is `Fill.Foreground`, `fontColor()` is `Font.Color` |
| `RichText.builder().run(...)`, `runs()` (RunInfo) | `new RichText(new TextRun(...) { Bold = true, ... })`, `Runs` (TextRun) | |
| `setComment(author, text)`, `addThreadedComment` returns id, `addThreadedReply(id, ...)` | `cell.Comment = new Comment(...)`, `ThreadedComments.Add` returns `ThreadedComment`, `.Reply(...)` | |
| `removeHyperlink()` | `cell.Hyperlink = null` | |
| `range.merge()`, `mergedRegions()` | `range.Merge()`, `sheet.MergedRegions` (`CellRange`) | |
| `freezeRow()` / `freezeColumn()` | `FrozenRows` / `FrozenColumns` | |
| `insertRows`, `setColumnWidth`, `setRowHeight`, `setColumnHidden`, `autoFitColumn`, `columnWidth`, `isRowHidden`, `rowOutlineLevel`, `groupRows`, `setRowGroupCollapsed`, `groupColumns` | `Rows.Insert`, `Columns[i].Width`, `Rows[i].Height`, `Columns[i].Hidden`, `Columns[i].AutoFit()`, `Rows[i].Hidden`, `Rows[i].OutlineLevel`, `Rows.Group`, `Rows.SetGroupCollapsed`, `Columns.Group` | |
| `sheet.subtotal(firstRow, lastRow, groupBy, hasHeader, fn, col)` | `sheet["A1:C5"].Subtotal(groupBy, fn, [col], hasHeaders)` | the range is given as a block; Java gave rows only |
| `SortKey.desc(1)`, `SortKey.asc(0)` | `new SortKey(1, Ascending: false)`, `new SortKey(0)` | |
| `setAutoFilter`, `setColumnFilter`, `setColumnCustomFilter`, `applyAutoFilter` | `AutoFilter.Range = `, `AutoFilter[c].FilterByValues`, `AutoFilter[c].FilterByCondition`, `AutoFilter.Apply()` | |
| `find(text).length`, `replaceAllRegex` | `FindAll(text).Count`, `ReplaceAll(..., new FindOptions { UseRegex = true })` | |
| `Validation.builder().type("whole").operator("between")...errorMessage(title, msg)`, `v.type().equals("list")` | `Validation.WholeNumber(ValidationOperator.Between, "1", "100") with { ErrorTitle, ErrorMessage }`, `v.Type == ValidationType.List` | strings became enums |
| `conditionalFormats().map(ConditionalFormat::type)` = `cellIs, colorScale, dataBar` | `ConditionalFormats.Select(c => c.Rule.GetType())` = `CellValueRule, ColorScaleRule, DataBarRule` | one record type per rule |
| `setTableStyle`, `resizeTable`, `setTableTotalsRow`, `setTableColumnTotals`, `convertTableToRange` (all return `true`) | `table.Style = `, `table.Range = `, `table.ShowTotalsRow = `, `Columns[n].TotalsFunction = `, `table.ConvertToRange()` (void; throw on failure) | Java asserted the `true` returns; .NET has none to assert. Design difference. |
| `Chart.column()...anchor(6, 0, 20, 8).build()` + `addChart`; `ChartView` | `sheet.Charts.Add(ChartType.Column, new DrawingAnchor(6, 0, 20, 8))` then `Title`, `Series.Add`, `Legend.Position`, axis titles; `Chart` | `type()` "column" / `grouping()` "clustered" are `ChartType.Column` / `ChartGrouping.Clustered` |
| `setSeriesFill`, `setSeriesDataLabels(0, DataLabels.value())`, `addTrendline(0, Trendline.linear())`, `seriesFill(0)`, `trendlines(0)` | `Series[0].FillColor`, `Series[0].DataLabels = new DataLabels { ShowValue = true }`, `Series[0].Trendlines.Add(new Trendline(TrendlineType.Linear))` | |
| `addSparkline("line", data, location)`, `type()` "line" | `Sparklines.Add(SparklineType.Line, data, location)`, `SparklineType.Line` | |
| `addPicture(bytes, "png", 1, 1, 6, 5)`, `pictureAnchor(0)` int[4], `pictureFormat` "png", `removePicture(0)` returns `true` | `Pictures.Add(bytes, PictureFormat.Png, new DrawingAnchor(1, 1, 6, 5))`, `Anchor.From/To`, `PictureFormat.Png`, `Pictures.RemoveAt(0)` (void) | the Java test builds the PNG with `BufferedImage`/`ImageIO`; the BCL has no portable imaging, so the port encodes the same 40x20 PNG by hand |
| `addShape("rect", text, r1, c1, r2, c2)`, `setShapeFillColor`, `shapeType`, `isTextBox` | `Shapes.Add("rect", anchor, text)` (returns `Shape`), `shape.FillColor`, `Geometry`, `IsTextBox` | |
| `addPivotTable` + `pivotTable(i)`, `addRowField`, `addDataField` | `PivotTables.Add(...)` returns `PivotTable`, `RowFields.Add`, `DataFields.Add` | |
| `renameSheet(i, n)`, `moveSheet`, `copySheet`, `deleteSheet`, `setActiveSheet(i)` | `sheet.Name = `, `Sheets.Move`, `Sheets.AddCopy`, `Sheets.RemoveAt`, `wb.ActiveSheet = sheet` | |
| `SheetProtection.builder()`, `addProtectedRange`, `protectStructure`, `isStructureProtected`, `unprotectStructure` | `new SheetProtection { ... }`, `ProtectedRanges.Add`, `wb.Protect(pw)`, `IsStructureLocked`, `wb.Unprotect(pw)` | |
| `saveEncrypted(path, pw)`, `open(path, pw)` | `Save(path, new SaveOptions { Password })`, `Open(path, new OpenOptions { Password })` | |
| `setLandscape`, `setPaperSize`, `setPageMargins(6 doubles)`, `fitToPage`, `addRowPageBreak`, `isLandscape` | `PageSetup.Orientation`, `PaperSize`, `Margins = new PageMargins(...)`, `FitToPages`, `AddRowBreak`, `PageOrientation.Landscape` | |
| `renderPdf(path, PdfOptions.builder()...)` | `SavePdf(path, new PdfOptions { SheetNames, FitTo = PdfFit.Width, PageRange })` | |
| `renderSheetPng/Svg(index, ...)`, `sheetPageCount(i)` | `sheet.RenderImage(ImageFormat.Png/Svg, ...)`, `sheet.GetPageCount()` | |
| `openBytes(csv, "csv")`, `textToColumns("A1:A2", ";")`, `CsvOptions.create().sheet(..).values(RAW)` | `Workbook.Open(bytes, WorkbookFormat.Csv)`, `sheet["A1:A2"].TextToColumns(";")`, `new CsvOptions { SheetName, Values = CsvValues.Raw }` | |
| `importData` with a Java record, columns `sku, name, price` | a C# record `Item(Sku, Name, Price)`, columns `Sku, Name, Price` | .NET property names are PascalCase |
| `StreamWorkbook.create(p)`, `StreamStyle.defaults().withBold()`, `row(r)` | `new StreamWorkbook(p)`, `StreamStyle.Default with { Bold = true }`, `StartRow(r)` | |
| `registerFunction` / `registerTypedFunction` | `Functions.RegisterNumeric` / `Functions.Register` | |
| `Functions.count/contains/list()` | `BuiltInFunctions.Count/Contains/Names` | |
| `setTitle`, `setCustomProperty`, `customProperties()` | `Title = `, `SetCustomProperty`, `CustomProperties` | |
| `wb.precedents(s, r, c)` / `dependents` as `CellRef.a1()` | `cell.Precedents` / `Dependents` as `Cell.Address` | |
| `setFormulaLocalized(f, locale)` | `SetLocalizedFormula(f, locale)` | |
| `setDefaultFont(defaultStyle().toBuilder()...)`, `Theme.builder().color(slot, c)` | `DefaultFont = DefaultStyle.Font with { ... }`, `new Theme { ... }.WithColor(slot, c)` | |
| `removeDuplicates`, `autoFill`, `consolidate` | same names on `CellRange` / `Sheet` | |
| `HtmlOptions.create().singleFile(true)`, `SaveOptions.create().formulaPolicy(..)` | `new HtmlOptions { SingleFile = true }`, `new SaveOptions { FormulaPolicy = .. }` | `IllegalArgumentException` became `ArgumentException` |
| `addSlicer(..., sheetIdx, r1, c1, r2, c2)`, `setSlicerSelection` (true), `slicerSelection`, `timelineRange` | `Slicers.Add(name, pivot, field, DrawingAnchor)`, `SelectedItems = [...]`, `Timelines.Add(..., DrawingAnchor, sheet)`, `Selection = new DateRange(...)` | the `true` returns have no counterpart |
| `FormControl.checkBox(name).text(..).linkedCell(..).build()` | `FormControls.Add(FormControlType.CheckBox, anchor, name)`, then `Text`, `LinkedCell`, `Checked`, ... | |
| `splitToFiles` returns `List<Path>` | `SplitToFiles` returns file path strings | |
| `Color.theme(1).toHex()` throws `IllegalStateException` | `Color.FromTheme(1).ToHex()` throws `InvalidOperationException` | per 16-design-conventions |
| `wb.sheets().indexOf(s)` | `wb.Sheets.ToList().IndexOf(s)` | `SheetCollection.IndexOf` takes a name; the list search keeps Java's meaning. Still -1: `Sheet` has reference equality in .NET too |
| `importData` of a bean with `getZeta()`/`getAlpha()` gives headers `alpha`, `zeta` | a class with properties `Zeta`, `Alpha` gives `Alpha`, `Zeta` | same documented order (by name); the header is the .NET property name, so its case differs. Design difference, not a defect |

## Observations (no test failure)

- The SDK source under `sdk-dotnet/src` is newer than the package in more places than issue 1: for example
  the source has `Picture.GetData()` where the package's XML lists a `Picture.Data` property. The tests
  compile against the package; nothing here uses those members.

## Lead follow-up (2026-11-17): FormulaToR1C1/FormulaToA1 anchor — fixed in source, not yet shipped
Re-ran the repro against the axyra-sheets source at HEAD `e5fab45d` ("fix(engine): the formula anchor measured off A1…")
via a ProjectReference to `sdk-dotnet/src/Axyra.Sheets` and `AXYRA_NATIVE_PATH=target/debug/libaxyra.dylib`:
`FormulaToR1C1("A1+$B$2",2,2)` → `R[-2]C[-2]+R2C2`, `FormulaToA1("R[-2]C[-2]+R2C2",2,2)` → `A1+$B$2`,
`FormulaToR1C1("SUM(A1:B2)",0,0)` → `SUM(RC:R[1]C[1])` — all as documented. So the bug is in package
`1.0.0-fl.20261116` only; drop the known-issue trait once the examples move to a build containing that commit.
