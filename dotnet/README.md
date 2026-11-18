# Axyra Sheets for .NET Examples

C# counterparts of the [Java examples](../java/), using properties, indexers,
records and collection APIs from the `Axyra.Sheets` NuGet package.

## Requirements

- .NET SDK 8 or later, plus .NET 8 runtime (projects target `net8.0`).
- Bash for `run-tests.sh` (macOS/Linux, Git Bash or WSL).
- Network access for NuGet restore; `nuget.config` includes the Axyra evaluation feed.

The native engine ships in the NuGet package. Java and Microsoft Excel are not
required. Run all commands below from this `dotnet/` directory.

```sh
dotnet build Axyra.Sheets.Examples.sln
dotnet run --project src/Axyra.Sheets.Examples -- list
dotnet run --project src/Axyra.Sheets.Examples -- tutorial/quick-start
dotnet run --project src/Axyra.Sheets.Examples -- feature-demo/insert-chart
bash run-tests.sh
```

Generated files go to `output/` unless an output directory is supplied after the
example name. `all` runs only the 20 documentation examples, not the 20 feature
demos. Run feature demos individually.

## Tests and evaluation mode

The package is pinned to **1.0.0-fl.20261116** in `Directory.Build.props`.
Evaluation output includes notice sheets/banners and a combined limit of 100
opens/saves/renders/stream creations per process. The test runner builds once
and launches a fresh process per test class; running all classes through one
`dotnet test` process can hit the evaluation limit.

```sh
bash run-tests.sh QuickStartIntegrationTest
bash run-tests.sh --known-issues
```

The normal suite has 130 tests: 129 pass and digital signing is skipped without
a license. The separate known-issue test currently fails: formula conversion
anchors in this package are one-based while newer SDK documentation describes
zero-based anchors. See [conversion notes](CONVERSION-NOTES/regression-features.md).
These results were measured on macOS arm64; rerun on your deployment platform.

## Documentation and migration

- [.NET documentation](https://doc.keikai.io/axyra/dotnet/)
- [Java examples](../java/)
- [API mapping and output comparisons](CONVERSION-NOTES/)

The implementation repository maintains its .NET API guides under
`axyra-sheets/docs/dotnet-api/`; the public documentation is maintained separately
in the sibling `doc` repository under `axyra/dotnet/`.
