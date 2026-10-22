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

Use try-with-resources to close each `Workbook` and release its native resources, as shown in the examples.
