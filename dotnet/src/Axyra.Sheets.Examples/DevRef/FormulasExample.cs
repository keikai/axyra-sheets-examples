using Axyra.Sheets.Calc;
using Axyra.Sheets.Formula;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class FormulasExample : IExample
{
    public string Name => "dev-ref/formulas";

    public string Description => "Calculate formulas, dynamic arrays, localized formulas, and a .NET UDF";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Formulas");
        workbook.CalcMode = CalcMode.Manual;
        sheet["A2:A6"].SetNumbers([3, 1, 3, 2, 1]);
        sheet[0, 0].Formula = "SORT(UNIQUE(A2:A6))";
        workbook.Functions.RegisterNumeric("TRIANGLE", args => args[0] * (args[0] + 1) / 2);
        sheet[0, 2].Formula = "TRIANGLE(10)";
        sheet[0, 3].SetLocalizedFormula("SUMME(A2:A6)", "de-DE");
        workbook.Recalculate();

        CellValue udf = sheet[0, 2].Value;
        CellValue sum = sheet.Evaluate("SUM(A2:A6)");
        Console.WriteLine($"Functions: {BuiltInFunctions.Count}; XLOOKUP: {BuiltInFunctions.Contains("XLOOKUP")}; "
            + $"UDF: {udf}; sum: {sum}");
        workbook.Save(Path.Combine(outputDirectory, "formulas.xlsx"));
    }
}
