using Axyra.Sheets.Calc;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class WorkbookAndSheetExample : IExample
{
    public string Name => "dev-ref/workbook-and-sheet";

    public string Description => "Manage sheets, workbook settings, page layout, and named ranges";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet summary = workbook.Sheets.Add("Summary");
        Sheet config = workbook.Sheets.Add("Config");
        config[1, 1].Value = 0.05;

        summary.TabColor = Color.FromRgb(0x2F, 0x55, 0x97);
        summary.PageSetup.Header = "&C&\"Calibri,Bold\"Quarterly Report";
        summary.PageSetup.Footer = "&LPrinted &D&RPage &P of &N";
        summary.Rows[0].Height = 28;
        summary.Columns[0].Width = 24;
        summary.FreezePanes(1, 0);
        config.Visibility = Visibility.VeryHidden;

        workbook.CalcMode = CalcMode.Manual;
        workbook.SetIterativeCalculation(true, 100, 0.001);
        workbook.Date1904 = false;
        workbook.Names.Add("TaxRate", "Config!$B$2");
        foreach (NamedRange name in workbook.Names)
        {
            Console.WriteLine(name.Name + " = " + name.Formula);
        }
        workbook.Save(Path.Combine(outputDirectory, "workbook-and-sheet.xlsx"));
    }
}
