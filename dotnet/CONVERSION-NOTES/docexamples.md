# Conversion notes: documentation examples (tutorial / guides / dev-ref)

Scope: `Tutorial/ReadWriteExample`, `Tutorial/RenderExample`, the 5 `Guides/*` and 12 `DevRef/*`
examples, and the tests `GuideExamplesIntegrationTest`, `RemainingExamplesIntegrationTest`,
`ExampleCatalogTest`. Package `Axyra.Sheets 1.0.0-fl.20261116`, evaluation mode (no license).

## How the port was checked (verified)

Each Java example was run from `axyra-sheets-examples/java/target/classes` (JDK 21, same eval build) and
each .NET example from the console app, into separate directories. Then:

- **Console output**: identical values for all 19 examples. Only the language's default
  formatting differs (Java `2479.0` / `Number[value=2479.0]` / `true` / `UNLICENSED`, .NET `2479` /
  `Number { Value = 2479 }` / `True` / `Unlicensed`), exactly as in the reference `QuickStartExample`.
- **Rendered and exported files**: all 14 non-XLSX outputs (PDF, PNG, SVG, JPEG, TIFF, HTML, CSV,
  JSON) are **byte-identical** between Java and .NET.
- **XLSX outputs**: 23 workbooks re-opened with the .NET SDK and dumped (sheets, visibility, names,
  calc settings, frozen panes, page setup, column widths, every value/formula/formatted text/style/
  comment/hyperlink/rich text, charts, tables, CF, validations, chart sheets). 22 are identical; the
  one difference is entry 1 below.
- Tests: `GuideExamplesIntegrationTest` 1/1 pass, `RemainingExamplesIntegrationTest` 1/1 pass,
  `ExampleCatalogTest.EveryExampleHasAUniqueNameAndDescription` passes; see entry 6 for the other one.

## 1. `Validation.List(...)` defaults `ShowInputMessage` to true (Java: false)

- Java: `Validation.list("APAC","EMEA","Americas")` sets `allowBlank`, `showErrorMessage`, no
  `showInputMessage` → the saved rule has `showInputMessage` false.
- .NET: `Validation.List("APAC","EMEA","Americas")` has `ShowInputMessage = true`
  (`Validation.cs`: `public bool ShowInputMessage { get; init; } = true;`).
- Repro (ContentObjectsExample): `sheet.Validations.Add("A2:A20", Validation.List("APAC","EMEA","Americas"));`
  save, reopen, read `sheet.Validations[0].Rule.ShowInputMessage`.
- Observed: Java file → `False`; .NET file → `True`. Everything else in the rule is equal.
- Classification: **intentional design difference** (documented in `docs/dotnet-api/09`: "`AllowBlank`,
  `ShowErrorMessage` and `ShowInputMessage` are on by default"). No visible effect here because the rule
  has no input title/message, but the saved XML differs between SDKs. The faithful C# code was kept.

## 2. `CellStyle` has no public JSON form (Java `toJson()` / `CellStyle.fromJson`)

- Java `StylesExample`: `CellStyle restored = CellStyle.fromJson(header.toJson());`
- .NET: `CellStyle.ToJson` / `FromJson` exist in the source but are `internal`; the package exposes no
  serialization of a style. `CellStyle` is a value record, so the port copies it (`header with { }`)
  and says so in a comment. The example's description was changed from "Build, serialize, register…"
  to "Build, copy, register…".
- Classification: **missing API / intentional design difference** (records are the .NET way to store
  and copy styles; `docs/dotnet-api/04` never mentions JSON). Verified: no `ToJson`/`FromJson` member on
  `CellStyle` in `Axyra.Sheets.xml`.

## 3. Chart sheets cannot be created from an existing chart definition

- Java `ChartsExample`: one immutable `Chart` is built, then `sheet.addChart(chart)` and
  `workbook.addChartSheet("Revenue", chart)` reuse it.
- .NET: charts are live objects; `sheet.Charts.Add(type, anchor)` and `wb.ChartSheets.Add(name, type)`
  each create an empty chart, which is then configured. The port applies the same settings to both
  through a `Configure(Chart)` helper. Output verified equal (chart sheet title "Revenue by Region",
  2 series, legend bottom, data labels on, white chart area).
- Classification: **intentional design difference** (live-object model, `docs/dotnet-api/16`).

## 4. `Color.Parse` returns `Color?`

- Java `Color.parse("#2F5597")` returns a `Color`. .NET `Color.Parse(string?)` returns `null` for text
  it cannot read, so passing it to `Fill`'s implicit conversion or `Borders.Around(style, Color)` gives
  nullable warning CS8604. The port uses `Color.Parse("#…")!`.
- Classification: **intentional design difference** (documented in `docs/dotnet-api/04`: "else null").
  Minor ergonomics: a typo in a hex literal surfaces later as an `ArgumentNullException`, not at Parse.

## 5. JSON imported through `System.Text.Json` arrives as text unless unwrapped

- Java `WriteDataToExcelExample`: Jackson `readValue(JSON, List<Map<String,Object>>)` yields `Double`
  for `1720.0`, and `importData` writes a number.
- .NET: the natural equivalent, `JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(json)`,
  gives `JsonElement` values. `Sheet.ImportData` does not recognise `JsonElement`, so it writes its
  `ToString()`.
- Repro:
  ```csharp
  var maps = JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(
      """[{"region":"Americas","product":"Widget","amount":1720.0},{"region":"APAC","amount":null}]""")!;
  sheet.ImportData(0, 0, maps, new ImportOptions { Columns = ["region", "product", "amount"] });
  Console.WriteLine(sheet[1, 2].Value);
  ```
  Observed: `Text { Value = 1720.0 }` (a number stored as text, silently).
- The port unwraps each `JsonElement` to `double` / `string` / `bool` / `null` first (`Plain(...)`), so
  the output matches Java (verified: same cells in `data-export.xlsx`).
- Classification: **doc gap** (ImportData's doc lists dictionaries but does not say values must be plain
  CLR values; `JsonElement` is the common way JSON reaches a dictionary in .NET). Possible enhancement:
  unwrap `JsonElement`/`JsonNode` in `DataImport.ToCellValue`. Whether the text coercion is intended is
  a guess.

## 6. Feature-demo catalog integration (resolved)

The 20 feature demos and both integration-test classes are now present.
The 2026-11-17 workspace verification ran `ExampleCatalogTest` successfully
(2 tests), including coverage of every registered feature demo. The earlier
compile-stub limitation no longer applies.

## Translation choices (not SDK findings)

- `LicenseAndNativeLoaderExample`: .NET has no JVM native loader. Mapping: Java `Workbook.licenseStatus()`
  → `License.Status`; `License.isFeatureEnabled` → `License.IsFeatureEnabled`; "native library loaded" is
  shown the same way, by creating a `Workbook` (its static constructor loads and checks the engine from
  `runtimes/<rid>/native/`). Java `-Daxyra.native.path` → env var `AXYRA_NATIVE_PATH`; a load failure is
  a `TypeInitializationException` wrapping `DllNotFoundException` / `AxyraException` (docs 01). License
  installation is `License.Apply(...)` (docs 15). These are described in the class comment; console
  output is the same three lines as Java.
- Name mapping: `sheet.cell(r,c).setValue` → `sheet[r,c].Value`; `setColumnWidth`/`setRowHeight`/`autoFitColumn`
  → `Columns[i].Width` / `Rows[i].Height` / `Columns[i].AutoFit()`; `setLandscape`/`setPaperSize`/
  `setPageMargins`/`setPrintTitleRows`/`setHeader`/`setFooter` → `sheet.PageSetup.*`;
  `renderPdf` → `SavePdf`; `renderSheetPng/Svg(index, opts)` and `Renderer.toJpeg/toTiff` →
  `sheet.RenderImage(ImageFormat.*, opts)`; `sheetPageCount(i)` → `sheet.GetPageCount()`;
  `workbook.evaluate(f, sheet)` → `sheet.Evaluate(f)`; `Functions.count/contains` → `BuiltInFunctions`;
  `registerFunction` (numeric lambda) → `wb.Functions.RegisterNumeric`; `setFormulaLocalized` →
  `SetLocalizedFormula`; `Workbook.createColumnar()` → `new Workbook(WorkbookStorage.Columnar)`;
  `StreamWorkbook.create(p)`/`row(i)` → `new StreamWorkbook(p)`/`StartRow(i)`; `openBytes(b,"xlsx")` →
  `Open(b, WorkbookFormat.Xlsx)`; `saveBytes("json")` → `ToByteArray(WorkbookFormat.Json)`;
  `protectStructure(pw)` → `wb.Protect(pw)`; `setItemFilter(field, items…)` →
  `pivot.Fields[field].VisibleItems`; `setShowDataAs(0, t)` → `DataFields[0].ShowDataAs`;
  `PivotStyle.of(n).withRowStripes()` → `new PivotStyle(n) { ShowRowStripes = true }`.
- Record property names are PascalCase (`Region`, `OrderDate`, `Q1`), so `ImportOptions.Columns` names
  them that way; Template Markers still use the Java-style lower-case paths (`${lines.region}`), which
  the .NET resolver maps to the capitalised property (verified: identical output).
- Descriptions adapted where they named Java: "a Java UDF" → "a .NET UDF", "Import Java records" →
  "Import .NET records"; plus the styles one (entry 2). `Name` strings and output file names are unchanged.
- Eval limit: `RemainingExamplesIntegrationTest` runs 14 examples in one process (~32 counted
  opens/saves/renders), `GuideExamplesIntegrationTest` 5 (~13). Same grouping as Java; well under 100.
