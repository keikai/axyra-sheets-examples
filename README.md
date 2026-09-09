# Axyra Sheets Examples

Runnable Java examples for the [Axyra Sheets documentation](https://doc.keikai.io/axyra).
The package and folder names follow the documentation sections so that readers can
quickly find the complete program behind a snippet.

## Requirements

- JDK 17 or later
- Maven 3.9 or later, or the included Maven wrapper
- Axyra Sheets `0.1.0-SNAPSHOT` installed in your local Maven repository while
  the SDK is under development

Install the SDK from a sibling checkout:

```bash
cd ../axyra-sheets/sdk-java
mvn install -DskipTests
```

Published Axyra Sheets jars bundle their native libraries. When using a thin jar
built from source, build the JNI library and pass its absolute path:

```bash
cd ../axyra-sheets
cargo build -p axyra-sheets-ffi-jni
cd ../axyra-sheets-examples
mvn test -Daxyra.native.path=../axyra-sheets/target/debug/libaxyra_jni.dylib
```

Use `libaxyra_jni.so` on Linux and `axyra_jni.dll` on Windows.

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

