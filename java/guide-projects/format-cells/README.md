# Format Excel cells in Java

Turn raw sales data into a business report using reusable styles, readable dates,
USD currency, percentages, borders, alignment, and a total row. Preserve existing
formatting when highlighting a row, and flag low margins with a conditional rule.

## Run

From this directory, run these in order:

```sh
mvn compile exec:java -Dexec.mainClass=FormatExcelReport
mvn exec:java -Dexec.mainClass=FormatExistingReport
mvn exec:java -Dexec.mainClass=RenderFormattingPreviews
```

| Program | Output |
| --- | --- |
| `FormatExcelReport` | `formatted-sales.xlsx`, with numeric revenue/margins and a formula total |
| `FormatExistingReport` | `formatted-sales-highlighted.xlsx`, preserving formats and adding a below-18% margin rule |
| `RenderFormattingPreviews` | `output/previews/` with raw, formatted, and highlighted PNGs and an unformatted workbook |

`FormatExcelReport` must run first. Run `FormatExistingReport` before generating
previews. No chart or formula project is required.

## Reference output

The reference workbooks are in `expected/`; full-page PNGs are in
`expected/figures/`. Surrounding blank whitespace can be cropped for an article.

Revenue `12500.5` displays as `$12,500.50`, and margin `0.185` as `18.5%`.
C5 retains `SUM(C2:C4)` and displays `$36,501.25`. The conditional rule shows
South's `15.2%` in bold red text while retaining its percentage format.
The raw-data preview uses a neutral grid and readable dates, with raw revenue
and margin values. All previews come from native Axyra rendering.

[Guide](https://doc.keikai.io/axyra/sheets/java/guides/apply-excel-styles-formatting-java)

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
