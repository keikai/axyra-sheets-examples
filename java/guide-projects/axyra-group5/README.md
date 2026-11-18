# Axyra guides — Group 5

Runnable examples for articles 19–21. Requires Java 17+ and Maven. The POM pins
Axyra Sheets 1.0.0-Beta-Eval and its public evaluation repository. No desktop Excel
installation is needed. Evaluation notices are retained in all workbook fixtures.

Run from this directory:

```sh
mvn compile exec:java -Dexec.mainClass=ModernFormulaBatch
mvn exec:java -Dexec.mainClass=CreateExcelCharts
mvn exec:java -Dexec.mainClass=FormatExcelReport
```

- Article 19 reads `input/sales-input.xlsx`, writes
  `output/sales-calculated.xlsx`, and checks results after reopening. Expected
  revenue is 445 before the quantity correction and 495 afterward.
- Article 20 writes and reopens `sales-chart.xlsx` in this directory.
- Article 21 writes and reopens `formatted-sales.xlsx` in this directory.

`expected/` contains generated reference outputs. Regenerate the synthetic input
with `mvn exec:java -Dexec.mainClass=CreateFormulaInput` when needed. It overwrites
only the example's input fixture, so keep a separate copy if you modify that file.

The input schema has exactly four sales rows and three product rows. Extend
the ranges and acceptance checks for other data sizes. The XLOOKUP fallback is
zero for demonstration; production pricing should define an unknown-SKU policy.

Verified on Java 17 against the pinned build. These checks cover this fixture's
values, formulas, chart data, and key formats; inspect desktop spreadsheet layout
or PDF output separately when those are part of your report contract.

Guides: https://doc.keikai.io/axyra/guides/
Installation: https://doc.keikai.io/axyra/dev-ref/Installation
Licensing: https://doc.keikai.io/axyra/dev-ref/License

## Reproduce article figures

```sh
mvn compile exec:java -Dexec.mainClass=RenderReviewFigures
```

The helper reopens the reference chart/report workbooks, adds the monthly and
pie-chart fixtures, renders PNG previews, and records numeric formatting and
recalculation evidence in `output/review-figures/`. It uses Java 17 with the
pinned Axyra dependency. `expected/chart-variants.xlsx` supplies the Monthly
sheet (Jan–Dec, revenue 1000–12000) used for the line-chart preview.

The previews are native Axyra output, not screenshots of Microsoft Excel.
Evaluation notices are retained. Before/after article illustrations use the
verified numeric evidence; they are labeled separately from workbook renders.
