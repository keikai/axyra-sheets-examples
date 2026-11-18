using Axyra.Sheets.Examples.DevRef;
using Axyra.Sheets.Examples.FeatureDemo;
using Axyra.Sheets.Examples.Guides;
using Axyra.Sheets.Examples.Tutorial;

namespace Axyra.Sheets.Examples;

public static class ExampleRunner
{
    private static readonly IReadOnlyList<IExample> DocumentationExamples =
    [
        new QuickStartExample(),
        new ReadWriteExample(),
        new RenderExample(),
        new CreateExcelFileExample(),
        new ReadExcelFileExample(),
        new ExcelToPdfExample(),
        new WriteDataToExcelExample(),
        new GenerateReportExample(),
        new WorkbookAndSheetExample(),
        new CellAndRangeExample(),
        new FormulasExample(),
        new StylesExample(),
        new ChartsExample(),
        new ContentObjectsExample(),
        new PivotTableExample(),
        new ImportExportExample(),
        new RenderingReferenceExample(),
        new StreamingExample(),
        new TemplateMarkersExample(),
        new LicenseAndNativeLoaderExample(),
    ];

    /// <summary>
    /// Runnable counterparts of the feature demos on the Axyra demo site. They run one at a
    /// time rather than by <c>all</c>: together they open, save and render more than the
    /// evaluation build allows in one process (100 operations).
    /// </summary>
    private static readonly IReadOnlyList<IExample> FeatureDemos =
    [
        new CreateWorkbookExample(),
        new OpenWorkbookExample(),
        new SaveExportExample(),
        new ManageWorksheetsExample(),
        new CellReadWriteExample(),
        new BulkRangeExample(),
        new FormulasRecalculationExample(),
        new ImportObjectsExample(),
        new CellStylingExample(),
        new NumberFormatsExample(),
        new RowsColumnsExample(),
        new MergeFreezeExample(),
        new FindReplaceExample(),
        new SortFilterExample(),
        new DataValidationExample(),
        new ConditionalFormatExample(),
        new CreateTableExample(),
        new InsertChartExample(),
        new ExportPdfExample(),
        new RenderImageExample(),
    ];

    public static IReadOnlyList<IExample> Examples { get; } = [.. DocumentationExamples, .. FeatureDemos];

    public static IReadOnlyList<IExample> Documentation => DocumentationExamples;

    public static int Main(string[] args)
    {
        if (args.Length == 0 || args[0] == "list")
        {
            PrintExamples();
            return 0;
        }

        string output = ExampleSupport.PrepareOutputDirectory(args.Length > 1 ? args[1] : "output");
        if (args[0] == "all")
        {
            foreach (IExample example in DocumentationExamples) Run(example, output);
            return 0;
        }

        IExample selected = Examples.FirstOrDefault(e => e.Name == args[0])
            ?? throw new ArgumentException(
                $"Unknown example '{args[0]}'. Run without arguments to list examples.");
        Run(selected, output);
        return 0;
    }

    private static void Run(IExample example, string output)
    {
        Console.WriteLine($"\n== {example.Name} ==");
        example.Run(output);
    }

    private static void PrintExamples()
    {
        Console.WriteLine("Axyra Sheets documentation examples:\n");
        foreach (IExample example in Examples.OrderBy(e => e.Name, StringComparer.Ordinal))
            Console.WriteLine($"  {example.Name,-38} {example.Description}");
        Console.WriteLine("\nRun one with: dotnet run --project src/Axyra.Sheets.Examples -- <name> [output-directory]");
    }
}
