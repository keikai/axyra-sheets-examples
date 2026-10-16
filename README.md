# Axyra Sheets Examples

Runnable Java examples for the [Axyra Sheets documentation](https://doc.keikai.io/axyra).
The package and folder names follow the documentation sections so that readers can
quickly find the complete program behind a snippet.

## Requirements

- JDK 17 or later
- Maven 3.9 or later, or the included Maven wrapper

Maven downloads Axyra Sheets `1.0.0-Beta.FL.20261013-Eval` from
`https://mavensync.zkoss.org/eval`, configured in `pom.xml`. The published JAR
bundles its native libraries; no engine source checkout, local installation, or
native-path override is required. This is an evaluation build: generated
files carry evaluation notices (including an extra first sheet named
"Axyra Evaluation Copy"), and one process can perform at most 100 opens, saves,
renders, and streaming workbook creations combined.

Compile the examples and run the tests before running an example:

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

Each example owns its workbook with try-with-resources. This is required because
`Workbook` wraps native memory that must be released promptly.
