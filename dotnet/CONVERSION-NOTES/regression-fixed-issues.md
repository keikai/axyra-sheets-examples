# regression-fixed-issues: FixedIssuesTest

Java source: `axyra-sheets-examples/java/src/test/java/io/keikai/axyra/examples/regression/FixedIssuesTest.java`
.NET port: `tests/Axyra.Sheets.Examples.Tests/Regression/FixedIssuesTest.cs` (57 `[Fact]`s, one per Java `@Test`,
same method names in PascalCase, Java `@DisplayName` kept as `DisplayName`).

Package: `Axyra.Sheets 1.0.0-fl.20261116` (evaluation mode).

## Result

`./run-tests.sh FixedIssuesTest`: **57 passed, 0 failed, 0 skipped** (about 1.3 s).
No test needed the `known-issue` trait, so `./run-tests.sh --known-issues FixedIssuesTest` runs nothing.
None of the 0930/1013/1021/1022 defects reproduces on the .NET SDK.

`IssueTestSupport.cs` was used unchanged; no bug found in it.

## Evaluation limit (100 opens/saves/renders per process)

The class fits in one process, as it did in one JVM. Counted from the code: about 44 counted
operations (saves 27, opens 15, PDF renders 2, stream workbook 1; `GetPageCount` is not counted).
The whole class ran in one process without hitting the limit.

## API mapping that changed the shape of a test (not defects)

Verified by compiling against the package and running the tests.

| Java | .NET | Note |
|---|---|---|
| `wb.copySheetFrom(src, i, name)` | `wb.Sheets.AddCopy(src.Sheets[i], name)` | |
| `sheet.setAutoFilter(r)`, `setColumnFilter(c, v)`, `applyAutoFilter()` | `sheet.AutoFilter.Range = r`, `sheet.AutoFilter[c].FilterByValues(v)`, `sheet.AutoFilter.Apply()` | |
| `sheet.insertRows`, `insertColumns`, `isRowHidden`, `autoFitColumn`, `columnWidth` | `sheet.Rows.Insert`, `sheet.Columns.Insert`, `Rows[i].Hidden`, `Columns[i].AutoFit()`, `Columns[i].Width` | |
| `wb.sheetPageCount(i)`, `addRowPageBreak`, `rowPageBreaks()` | `sheet.GetPageCount()`, `sheet.PageSetup.AddRowBreak`, `PageSetup.GetRowBreaks()` | |
| `wb.evaluate(f, sheet)` | `sheet.Evaluate(f)` | |
| `wb.addTimeline(name, pivotName, field, sheetIdx, r1, c1, r2, c2)` | `wb.Timelines.Add(name, PivotTable, field, DrawingAnchor, Sheet)` | takes the pivot object, not its name |
| `boolean wb.setTimelineRange(name, start, end)` / `LocalDate[] timelineRange(name)` | `wb.Timelines[name].Selection = new DateRange(start, end)` (void) / `.Selection` | Java asserted the `true` return; .NET has no return value to assert, so that one assertion has no counterpart. Design difference (property setter). |
| `Chart.column()...showDataLabels(true).anchor(...).build()` + `addChart` | `sheet.Charts.Add(ChartType, DrawingAnchor)` then `Title`, `Series.Add`, `ShowDataLabels` | |
| `Chart.bar().barShape("cylinder")` | `chart.BarShape = BarShape.Cylinder` | see observation 1 |
| `addTextBox(text, r1,c1,r2,c2)` returns index; `isTextBox(i)` | `Shapes.AddTextBox(anchor, text)` returns `Shape`; `Shape.IsTextBox` | |
| `pictureAnchor(i)` / `chart.anchor()` / `shapeAnchor(i)` as `int[4]` | `DrawingAnchor` (`From.Row`, `From.Column`, `To.Row`, `To.Column`) | |
| `ConditionalFormat.builder().containsText(..).fillColor(..)` etc. | `new TextRule(TextRuleOperator.Contains, ..) { Style = new ConditionalStyle { FillColor = .. } }`, `CellValueRule`, `DataBarRule` | |
| `ConditionalFormat.builder().colorScale(Color...)` | `new ColorScaleRule(IReadOnlyList<ColorScalePoint>)` | .NET needs a threshold per colour; the test gives min / percentiles / max. Both throw `ArgumentException` (Java `IllegalArgumentException`) for 1 and 4 colours. |
| `Theme.builder().color(slot, c).majorFont(f).build()` | `new Theme { MajorFont = f }.WithColor(slot, c)` | |
| `wb.saveEncrypted(p, pw)` | `wb.Save(p, new SaveOptions { Password = pw })` | |
| `Renderer.toPdf(wb, opts)` | `wb.ToPdf(opts)` | |
| `wb.saveBytes("xlsb")` / `Workbook.openBytes(b, "xlsb")` | `wb.ToByteArray(WorkbookFormat.Xlsb)` / `Workbook.Open(b, WorkbookFormat.Xlsb)` | |
| `StreamWorkbook.create(p)`, `sw.row(i)` | `new StreamWorkbook(p)`, `sw.StartRow(i)` | |
| `wb.registerFunction(name, double[] -> double)` | `wb.Functions.RegisterNumeric` | |
| `License.status().isLicensed()` + `assumeFalse` | `License.Status.IsLicensed` + `Assert.SkipWhen` | |
| `CellStyle.fromJson(style.toJson())` (test #60) | **no JSON form of `CellStyle` in .NET** | Missing API / design difference: `CellStyle` is a value record, so equality is by construction. The port compares two independently built records with the same settings. The JSON round trip itself is not exercised. |
| Java `assertThrows(X.class, ..)` (accepts subclasses) | `Assert.ThrowsAny<X>` | same semantics |

## Observations (tests pass; recorded for the lead)

1. **#53 `BarShape` on a 2-D bar chart.** `chart.BarShape = BarShape.Cylinder` on a `ChartType.Bar` chart
   (not 3-D) does not throw, and reads back `Box` both in memory and after reopening, so the setting is
   silently dropped. The test only requires in-memory == reopened, so it passes. The XML doc says the
   property is "the solid a 3-D bar or column is drawn as", so this is consistent with the docs, but a
   silent no-op rather than a refusal. Verified with a standalone probe (scratch, not kept):
   ```csharp
   var c = s.Charts.Add(ChartType.Bar, new DrawingAnchor(0, 2, 10, 8));
   c.BarShape = BarShape.Cylinder;          // no exception
   Console.WriteLine(c.BarShape);           // Box (also Box after save + reopen)
   ```
   Java behaviour on this build was not run; its test accepts either outcome.
2. **Refusal branches taken on .NET** (the Java tests accept "works" or "refuses"; probe output):
   - #44 `.xlsb`: saved and reopened; comment and hyperlink survive (assertion path exercised).
   - #50 encrypted `.ods`: refused, `AxyraException: an encrypted save writes .xlsx, .xlsm, .xltx, .xltm or .xls; cannot encrypt ….ods`
     (early return; the ZIP-magic assertion is not reached).
   - #52 PDF `AllowCopy/AllowPrint = false` without password: refused, `AxyraException: pdf render: Invalid export options: PDF permissions need a password …`
     (early return).
3. **#22 thread safety** passes: the second thread gets an exception whose message contains
   "workbook is already in use", matching the .NET design doc ("Two threads using it at once get an exception").
