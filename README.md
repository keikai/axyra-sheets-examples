# Axyra Sheets Examples

Axyra Sheets is a server-side Java Excel library for enterprise spreadsheet processing, with modern formulas, charts, pivot tables, and PDF/image rendering. Its AI-friendly API and Axyra Agent Skill help teams build spreadsheet workflows faster.

[Website](https://keikai.io/axyra) ． [Quick Start](https://doc.keikai.io/axyra/quick-start/)

This repository contains runnable Java examples for the [Axyra Sheets documentation](https://doc.keikai.io/axyra). Package and folder names follow the documentation sections, making it easy to find the complete program behind each snippet.

## Requirements

- JDK 17 or later
- Maven 3.9 or later, or the included Maven wrapper

Maven automatically downloads Axyra Sheets and its required native libraries as configured in pom.xml. The included [evaluation version](https://doc.keikai.io/axyra/dev-ref/License) adds evaluation notices to generated files.

Build and test the examples:

```bash
./mvnw test
```

## Run examples

List the available examples:

```bash
./mvnw exec:java
```

Run one example. Generated files go to `output/` by default:

```bash
./mvnw exec:java -Dexec.args="tutorial/quick-start"
./mvnw exec:java -Dexec.args="guides/create-excel-file output"
```

Run every self-contained example:

```bash
./mvnw exec:java -Dexec.args="all"
```

`all` covers the documentation examples. The `feature-demo/*` examples, which
mirror the feature demos on the Axyra demo site, run one at a time: together
they open, save and render more files than the evaluation build allows in one
process.

```bash
./mvnw exec:java -Dexec.args="feature-demo/insert-chart"
```

### Windows PowerShell

The commands above are for bash, zsh, Git Bash and WSL. PowerShell splits an
unquoted `-Dexec.args=...` at the first dot, and Maven then fails with
`Unknown lifecycle phase ".args=..."`. Quote the whole `-D` argument and call
the wrapper as `.\mvnw`:

```powershell
.\mvnw test
.\mvnw exec:java
.\mvnw exec:java "-Dexec.args=tutorial/quick-start"
.\mvnw exec:java "-Dexec.args=guides/create-excel-file output"
.\mvnw exec:java "-Dexec.args=all"
.\mvnw exec:java "-Dexec.args=feature-demo/insert-chart"
```

In Command Prompt (`cmd.exe`), the bash form works with `.\mvnw` in place of
`./mvnw`, for example `.\mvnw exec:java -Dexec.args="tutorial/quick-start"`.

## Documentation mapping

| Documentation section | Example name | Source package |
|---|---|---|
| Tutorial: Quick Start | `tutorial/quick-start` | `examples.tutorial` |
| Tutorial: Read and Write | `tutorial/read-write` | `examples.tutorial` |
| Tutorial: Render | `tutorial/render` | `examples.tutorial` |
| Guide: Create an Excel File | `guides/create-excel-file` | `examples.guides` |
| Guide: Read an Excel File | `guides/read-excel-file` | `examples.guides` |
| Guide: Convert Excel to PDF | `guides/convert-excel-to-pdf` | `examples.guides` |
| Guide: Write Data to Excel | `guides/write-data-to-excel` | `examples.guides` |
| Guide: Generate Reports | `guides/generate-report` | `examples.guides` |
| Developer Reference | `dev-ref/*` | `examples.devref` |
| Feature demo: Create a new workbook | `feature-demo/create-workbook` | `examples.featuredemo` |
| Feature demo: Open an existing workbook | `feature-demo/open-workbook` | `examples.featuredemo` |
| Feature demo: Save and export to file | `feature-demo/save-export` | `examples.featuredemo` |
| Feature demo: Manage worksheets | `feature-demo/manage-worksheets` | `examples.featuredemo` |
| Feature demo: Write and read a cell | `feature-demo/cell-read-write` | `examples.featuredemo` |
| Feature demo: Bulk range read and write | `feature-demo/bulk-range` | `examples.featuredemo` |
| Feature demo: Formulas and recalculation | `feature-demo/formulas-recalculation` | `examples.featuredemo` |
| Feature demo: Import Java objects into a sheet | `feature-demo/import-objects` | `examples.featuredemo` |
| Feature demo: Cell styling | `feature-demo/cell-styling` | `examples.featuredemo` |
| Feature demo: Number and date formats | `feature-demo/number-formats` | `examples.featuredemo` |
| Feature demo: Rows and columns | `feature-demo/rows-columns` | `examples.featuredemo` |
| Feature demo: Merge cells and freeze panes | `feature-demo/merge-freeze` | `examples.featuredemo` |
| Feature demo: Find and replace | `feature-demo/find-replace` | `examples.featuredemo` |
| Feature demo: Sort and filter | `feature-demo/sort-filter` | `examples.featuredemo` |
| Feature demo: Data validation | `feature-demo/data-validation` | `examples.featuredemo` |
| Feature demo: Conditional formatting | `feature-demo/conditional-format` | `examples.featuredemo` |
| Feature demo: Create a table | `feature-demo/create-table` | `examples.featuredemo` |
| Feature demo: Insert a chart | `feature-demo/insert-chart` | `examples.featuredemo` |
| Feature demo: Export to PDF | `feature-demo/export-pdf` | `examples.featuredemo` |
| Feature demo: Render to an image | `feature-demo/render-image` | `examples.featuredemo` |

Use try-with-resources to close each `Workbook` and release its native resources, as shown in the examples.
