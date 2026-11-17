# Fill a sales template from Java records

From this directory, with Java 17+ and Maven:

```sh
mvn compile exec:java -Dexec.mainClass=FillReportTemplate
mvn exec:java -Dexec.mainClass=TemplateCases
```

The first command opens `input/sales-template.xlsx`, expands the Invoice detail
row into three lines, and writes `output/acme-sales.xlsx`. It reopens the file to
check customer and period, text SKUs, products, quantities, unit prices, formulas,
D/E number formats, the total label and value of 394, and remaining markers. The second checks a one-line report and rejects an empty list before
processing. Reference output is in `expected/`.

`CreateReportTemplate` regenerates the supplied input when explicitly run:

```sh
mvn exec:java -Dexec.mainClass=CreateReportTemplate
```

The total marker aggregates Java record amounts into a numeric value; it is not
an editable SUM formula. Define a rounding rule for financial production data.
Inspect layout in your target spreadsheet application. These examples do not
generate or validate PDF output. Evaluation marks remain intact.

The article prompt is in `prompt.txt`; it specifies a development workflow and
acceptance checks, not an archived generation transcript. Actual reference-run
commands and output are in `expected/verification.txt`.

After generating the workbooks, create article previews with:

```sh
mvn compile exec:java -Dexec.mainClass=RenderGuideFigures
```

The PNGs in `expected/` are rendered from the real workbooks using Axyra;
evaluation notices remain visible. The article versions trim outer whitespace.
These previews do not certify desktop Excel visual equivalence.
The template preview uses landscape orientation in memory so all five columns
are visible; it does not resave the workbook.
