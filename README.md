# Axyra Sheets Examples

Runnable Java and C#/.NET examples for Axyra Sheets, the spreadsheet engine for
reading, writing, calculating and rendering Excel files without Microsoft Office.

| SDK | Project | Requirements | Documentation |
|---|---|---|---|
| Java | [java/](java/) | JDK 17+, included Maven wrapper | [Java docs](https://doc.keikai.io/axyra/) |
| C# / .NET | [dotnet/](dotnet/) | .NET SDK 8+ and .NET 8 runtime | [.NET docs](https://doc.keikai.io/axyra/dotnet/) |

Each SDK has 40 examples with matching names: 3 tutorials, 5 guides,
12 developer-reference examples and 20 feature demos. Build and run from the
language directory so relative input and output paths resolve correctly.

## Java

```sh
cd java
./mvnw test
./mvnw exec:java -Dexec.args="tutorial/quick-start"
```

## .NET

```sh
cd dotnet
bash run-tests.sh
dotnet run --project src/Axyra.Sheets.Examples -- tutorial/quick-start
```

Run these blocks separately from the repository root. On Windows, use the Java
wrapper `mvnw.cmd`; the .NET test runner requires Bash (Git Bash or WSL).
See each language's README for prerequisites, example names and known issues.
The `all` runner executes the 20 documentation examples; run feature demos
individually to stay within the evaluation build's per-process limit.

## Repository layout

- `java/`: Maven project, resources, tests and standalone article projects in `guide-projects/`.
- `dotnet/`: .NET solution, embedded resources, tests and Java-to-.NET conversion notes.

Java's Maven version and .NET's NuGet version are selected independently. The
public examples use evaluation packages, whose output keeps evaluation notices.
A feature's presence in one SDK does not guarantee identical API shape or output
in the other; see [conversion notes](dotnet/CONVERSION-NOTES/).

The Java project previously lived at the repository root. Java source links now
start with `java/src/`; article projects start with `java/guide-projects/`.
Each language retains its own resources so it can be used independently.
