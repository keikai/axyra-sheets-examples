# Java Excel report examples

Create editable Excel charts and apply reusable styles to a sales report with
Axyra Sheets. Requires Java 17+ and Maven; no desktop Excel installation is needed.

Run the examples from this directory:

```sh
mvn compile exec:java -Dexec.mainClass=CreateExcelCharts
mvn exec:java -Dexec.mainClass=UpdateExcelChart
mvn exec:java -Dexec.mainClass=ExportExcelChart
mvn exec:java -Dexec.mainClass=FormatExcelReport
mvn exec:java -Dexec.mainClass=FormatExistingReport
mvn exec:java -Dexec.mainClass=RenderReportPreviews
```

| Program | Result |
| --- | --- |
| `CreateExcelCharts` | `sales-chart.xlsx`: four regions, actual and target revenue, one editable clustered column chart |
| `UpdateExcelChart` | `sales-chart-updated.xlsx`: adds Central and extends both series to row 6 |
| `ExportExcelChart` | `sales-chart.png`: chart only; `sales-chart.pdf`: one landscape page with the source table and chart |
| `FormatExcelReport` | `formatted-sales.xlsx`: dates, USD revenue, percentages, header and total styles |
| `FormatExistingReport` | `formatted-sales-highlighted.xlsx`: preserves row formats, applies a named currency style, and adds a margin-below-18% rule |
| `RenderReportPreviews` | `output/previews/`: line and pie chart variants, PNG previews, and an unformatted sales workbook |

Run `CreateExcelCharts` before the update and export examples. Run
`CreateExcelCharts`, `FormatExcelReport`, and `FormatExistingReport` before
`RenderReportPreviews`.
The formatting extension requires `formatted-sales.xlsx`.

## What to inspect

The main chart binds Actual to `Sales!$B$2:$B$5`, Target to
`Sales!$C$2:$C$5`, and both category ranges to `Sales!$A$2:$A$5`.
The update example extends all references through row 6. Fixed series ranges
must be extended when data grows. The PNG keeps data labels; the smaller PDF
chart omits them for readability.

The formatted report keeps revenue `12500.5` and margin `0.185` numeric while
displaying `$12,500.50` and `18.5%`. Cell C5 contains `SUM(C2:C4)` and displays
`$36,501.25`. The extension keeps the date/currency/percentage formats and
adds an Excel conditional-formatting rule that shows the South margin of
`15.2%` in bold red text. The percentage format remains in place.

`expected/` contains reference workbooks and exports. `expected/figures/`
contains full-page formatting renders and standalone chart images. For an
article illustration, surrounding blank page margins may be cropped. All
figures come from Axyra rendering. The raw-data preview has a neutral grid and
readable date formatting; revenue and margin retain their raw numeric display.
The monthly chart uses illustrative revenue with seasonal variation, from
8,600 to 17,800 USD. `RenderReportPreviews` also renders the saved conditional
format in `output/previews/highlighted-report.png`.

The chart and formatting reference files were generated with Axyra Sheets
`1.0.0-Beta` and an installed license. Without a license, normal example runs
use evaluation mode and include its notices. The programs leave license
installation to your application's startup code; see [Licensing and evaluation](https://doc.keikai.io/axyra/dev-ref/License).

## Dependency

The default POM uses `1.0.0-Beta-Eval` from the public evaluation repository.
The optional `ee` Maven profile selects `1.0.0-Beta` from the authenticated
repository. Configure the `keikai-ee` server in Maven settings when using it;
`axyra.ee.repo.id` can select a different existing server ID.

## Formula batch example

The project also retains a separate modern-formula example:

```sh
mvn exec:java -Dexec.mainClass=ModernFormulaBatch
```

It reads `input/sales-input.xlsx` and writes `output/sales-calculated.xlsx`.
Revenue is 445 before a quantity correction and 495 afterward. The fixture
has four sales rows and three product rows; adapt ranges for other data sizes.
The XLOOKUP zero fallback is demonstrative; define your own unknown-SKU policy.
Use `CreateFormulaInput` to regenerate the input fixture. Existing formula
reference fixtures retain their evaluation notices.

`RenderReviewFigures` is the earlier combined formula/chart evidence helper.
For the chart and formatting articles, use `RenderReportPreviews` above.

[Installation](https://doc.keikai.io/axyra/dev-ref/Installation) ·
[Charts](https://doc.keikai.io/axyra/dev-ref/Charts) ·
[Styles and formats](https://doc.keikai.io/axyra/dev-ref/Styles)
