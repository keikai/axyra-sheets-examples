# Article 23 — Migrate a POI sales report

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

Remaining differences on Axyra 1.0.0-Beta-Eval:

- Evaluation notice content is retained.
- POI's MINOR font scheme becomes NONE for the default font and the body cells
  A2:D3, A4, and C4:D4, despite requesting the scheme. Explicit Calibri 11,
  indexed color 8, and the other checked font properties match. Theme-dependent
  font behavior therefore remains a migration gap. The comparator reports each
  affected location separately rather than treating the schemes as equal.

ZIP/package metadata, desktop Excel appearance, and PDF output are not compared.
There are no chart, table, macro, external-link, or PDF migration scenarios in
this fixture. A passing fixture comparison is not a general compatibility claim.

POI 5.3.0 is kept intentionally for the reproducible baseline and independent
reader. The Axyra implementation itself imports no POI classes. Review remaining
usages before removing POI from any production module.
