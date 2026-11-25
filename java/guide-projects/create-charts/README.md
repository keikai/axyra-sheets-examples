# Create Excel charts in Java

Create an editable clustered column chart comparing actual and target revenue
for four regions, update its data, and export a chart PNG or a report PDF.

## Run

From this directory, run these in order:

```sh
mvn compile exec:java -Dexec.mainClass=CreateExcelCharts
mvn exec:java -Dexec.mainClass=UpdateExcelChart
mvn exec:java -Dexec.mainClass=ExportExcelChart
mvn exec:java -Dexec.mainClass=ExportChartFormats
mvn exec:java -Dexec.mainClass=RenderChartPreviews
```

| Program | Output |
| --- | --- |
| `CreateExcelCharts` | `sales-chart.xlsx`, with actual and target series and their worksheet references |
| `UpdateExcelChart` | `sales-chart-updated.xlsx`, adding Central and extending both series to row 6 |
| `ExportExcelChart` | `sales-chart.png` and a single landscape `sales-chart.pdf` with the data and chart |
| `ExportChartFormats` | `output/chart-export/`: a 1200×700 chart PNG, a chart-only PDF, and a report PDF |
| `RenderChartPreviews` | `output/previews/chart-variants.xlsx` plus column, line, and pie PNGs |

`CreateExcelCharts` must run first; the other programs open `sales-chart.xlsx`.
All chart examples are self-contained in this project.

## Reference output

`expected/` contains the editable workbooks, PNG, and PDF. `expected/figures/`
contains standalone column, line, and pie chart previews.

The main chart uses `Sales!$A$2:$A$5` as categories, `Sales!$B$2:$B$5` for
Actual, and `Sales!$C$2:$C$5` for Target. Update both references when adding rows.
The PNG retains data labels; the smaller PDF chart omits them for readability.
The monthly line chart uses illustrative revenue between USD 8,600 and 17,800.

[Guide](https://doc.keikai.io/axyra/sheets/java/guides/create-excel-charts-java)

## Dependency and licensing

Requires Java 17+ and Maven. The default POM uses Axyra Sheets
`1.0.0-Beta-Eval` from the public evaluation repository. The optional `ee`
profile selects `1.0.0-Beta` from the authenticated repository; configure the
`keikai-ee` Maven server or select an existing server with `axyra.ee.repo.id`.

The programs leave license installation to your application's startup code.
Without an installed license they run in evaluation mode, including its notices.
Reference outputs in `expected/` were generated with `1.0.0-Beta` and an installed
license. Microsoft Excel is not required to run the examples.

[Installation](https://doc.keikai.io/axyra/sheets/java/dev-ref/Installation) ·
[Licensing and evaluation](https://doc.keikai.io/axyra/sheets/java/dev-ref/License)

## Export an existing chart to PNG and PDF

`ExportChartFormats` opens the Sales chart created by `CreateExcelCharts`.
It writes `sales-chart.png`, `sales-chart-only.pdf`, and `sales-report.pdf` to
`output/chart-export/`. For PDF output, the chart is moved below the source
table to `A8:H24`. The chart-only PDF uses print area `A8:H24`; the report
uses `A1:H24` to include the source table. Both PDF layouts omit data labels for
readability. The saved source workbook is not overwritten.

The chart-only PDF is a portrait page containing the chart, with the source
table excluded. The PDF page edges are not cropped to the chart bounds.

To supply another workbook with the same Sales sheet and chart layout:

```sh
mvn exec:java -Dexec.mainClass=ExportChartFormats -Dexec.args="sales-chart.xlsx output/custom-export"
```

Reference files for these exports are under `expected/chart-export/`.
