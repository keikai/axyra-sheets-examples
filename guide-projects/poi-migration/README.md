# Migrate a POI sales report

From this directory, with Java 17+ and Maven:

```sh
mvn compile exec:java -Dexec.mainClass=PoiSalesReport
mvn exec:java -Dexec.mainClass=AxyraSalesReport
mvn exec:java -Dexec.mainClass=CompareReports
```

Outputs are `output/poi-sales.xlsx` and `output/axyra-sales.xlsx`. The independent
POI reader checks authored cells A1:D4, types, formulas and cached values,
number formats, fonts, fill patterns, alignment, wrapping, border styles,
defaults, widths, freeze panes, merges, and tables. Reference files and the
comparison transcript are in `expected/`.

The fixture uses explicit Calibri 11 fonts without theme-font associations.
Font schemes are checked strictly; every saved font must have scheme NONE.
The only expected difference in this comparison is evaluation notice content,
which is retained. A passing comparison verifies the selected workbook properties
and authored cells in this example, not every Apache POI workbook.

ZIP/package metadata, desktop Excel appearance, and PDF output are not compared.
There are no chart, table, macro, external-link, or PDF migration scenarios in
this fixture. A passing fixture comparison is not a general compatibility claim.

POI 5.3.0 is kept intentionally for the reproducible baseline and independent
reader. The Axyra implementation itself imports no POI classes. Review remaining
usages before removing POI from any production module.

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
The evaluation renderer also adds a banner when previewing the POI workbook.
