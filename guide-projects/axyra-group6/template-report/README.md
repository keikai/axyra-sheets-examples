# Article 22 — Fill a sales template from Java records

From this directory, with Java 17+ and Maven:

```sh
mvn compile exec:java -Dexec.mainClass=FillReportTemplate
mvn exec:java -Dexec.mainClass=TemplateCases
```

The first command opens `input/sales-template.xlsx`, expands the Invoice detail
row into three lines, and writes `output/acme-sales.xlsx`. It reopens the file to
check text SKUs, formulas, amounts, money formats, the total of 394, and remaining
markers. The second checks a one-line report and rejects an empty list before
processing. Reference output is in `expected/`.

`CreateReportTemplate` regenerates the supplied input when explicitly run:

```sh
mvn exec:java -Dexec.mainClass=CreateReportTemplate
```

The total marker aggregates Java record amounts into a numeric value; it is not
an editable SUM formula. Define a rounding rule for financial production data.
Inspect layout in your target spreadsheet application. These examples do not
generate or validate PDF output. Evaluation marks remain intact.
