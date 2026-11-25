# Calculate Excel formulas in Java

Use XLOOKUP to find prices, FILTER to return matching orders, and UNIQUE/SORT
to list regions. Changing a quantity automatically updates revenue from 445 to
495. The XLSX retains its formulas and calculated results.

## Run

From this directory:

```sh
mvn compile exec:java -Dexec.mainClass=ModernFormulaBatch
```

Input: `input/sales-input.xlsx`. Output: `output/sales-calculated.xlsx`.
Compare it with `expected/sales-calculated.xlsx`.

To use another workbook with the same layout:

```sh
mvn exec:java -Dexec.mainClass=ModernFormulaBatch -Dexec.args="input/sales-input.xlsx output/custom-report.xlsx"
```

`CreateFormulaInput` regenerates the synthetic input workbook. It overwrites
`input/sales-input.xlsx`, so keep a separate copy of any changes first.
`RecordFormulaResults` runs after `ModernFormulaBatch` and writes the before/after
totals and a before-correction workbook into `output/results/`:

```sh
mvn exec:java -Dexec.mainClass=RecordFormulaResults
```

## Adapt the example

The sample contains four sales rows and three product rows. Adjust ranges to
your record count and define how unknown SKUs should be handled. The example
uses zero as the XLOOKUP fallback. A dynamic-array formula needs one anchor
and an empty spill area.

[Guide](https://doc.keikai.io/axyra/sheets/java/guides/calculate-modern-excel-formulas-java)

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
