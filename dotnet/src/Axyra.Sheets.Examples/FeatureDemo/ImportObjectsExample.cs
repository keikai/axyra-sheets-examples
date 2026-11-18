using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Import .NET objects straight onto a sheet.</summary>
/// <remarks>
/// <para><c>ImportData</c> reads a collection of objects or records and writes it as a block —
/// no per-cell loop, no manual header row. <see cref="ImportOptions"/> picks which properties go
/// in, in what order, under what labels, and how the header is styled.</para>
/// <para>A dictionary is a row too, keyed by column name: <see cref="Maps"/> imports records that
/// have no class behind them with the same call.</para>
/// </remarks>
public sealed class ImportObjectsExample : IExample
{
    public string Name => "feature-demo/import-objects";

    public string Description => "Import a list of beans or maps onto a sheet with importData";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Beans())
        {
            wb.Save(Path.Combine(outputDirectory, "import-objects-beans.xlsx"));
        }
        using (Workbook wb = ChosenColumns())
        {
            wb.Save(Path.Combine(outputDirectory, "import-objects-chosen-columns.xlsx"));
        }
        using (Workbook wb = Maps())
        {
            wb.Save(Path.Combine(outputDirectory, "import-objects-maps.xlsx"));
        }
        using (Workbook wb = ImportAndTotal())
        {
            wb.Save(Path.Combine(outputDirectory, "import-objects-total.xlsx"));
        }
        Console.WriteLine("Wrote four object-import workbooks to " + outputDirectory);
    }

    /// <summary>All properties, a styled header row, one call.</summary>
    private static Workbook Beans()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Trips");

        int rows = sheet.ImportData(0, 0, Trips(), new ImportOptions
        {
            Headers = true,
            ColumnLabels = ["City", "Nights", "Cost"],
            HeaderStyle = new CellStyle
            {
                Font = new Font { Bold = true },
                Fill = Color.FromRgb(0xDD, 0xEB, 0xF7),
            },
        });

        sheet[rows + 2, 0].Value = "ImportData wrote " + rows + " rows from a List<Trip>.";
        Widen(sheet);
        return wb;
    }

    /// <summary>A subset, in an order of your choosing, under your own labels.</summary>
    private static Workbook ChosenColumns()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Costs");

        sheet.ImportData(0, 0, Trips(), new ImportOptions
        {
            Columns = ["Cost", "City"],
            Headers = true,
            ColumnLabels = ["Cost", "Destination"],
            HeaderStyle = new CellStyle { Font = new Font { Bold = true } },
        });

        sheet[Trips().Count + 2, 0].Value = "Columns = [\"Cost\", \"City\"] reordered and dropped a property.";
        Widen(sheet);
        return wb;
    }

    /// <summary>Keyed rows with no class behind them: each dictionary is a row, its keys the columns.</summary>
    private static Workbook Maps()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Dynamic");

        List<Dictionary<string, object>> rows =
        [
            Row("alpha", 12, "open"),
            Row("beta", 47, "closed"),
            Row("gamma", 8, "open"),
        ];

        // The keys' insertion order fixes the column order; numbers stay numeric.
        sheet.ImportData(0, 0, rows, new ImportOptions
        {
            Headers = true,
            HeaderStyle = new CellStyle { Font = new Font { Bold = true } },
        });

        sheet[rows.Count + 2, 0].Value = "Keys became the header row, in insertion order.";
        Widen(sheet);
        return wb;
    }

    /// <summary>The imported block is ordinary cells, so formulas apply on top of it.</summary>
    private static Workbook ImportAndTotal()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Trips");

        List<Trip> trips = Trips();
        int rows = sheet.ImportData(0, 0, trips, new ImportOptions
        {
            Headers = true,
            ColumnLabels = ["City", "Nights", "Cost"],
            HeaderStyle = new CellStyle { Font = new Font { Bold = true } },
        });

        int totalRow = rows + 1;
        sheet[totalRow, 0].Value = "Total";
        sheet[totalRow, 1].Formula = "SUM(B2:B" + rows + ")";
        sheet[totalRow, 2].Formula = "SUM(C2:C" + rows + ")";
        sheet[totalRow + 1, 0].Value = "Cost per night";
        sheet[totalRow + 1, 2].Formula = "C" + (totalRow + 1) + "/B" + (totalRow + 1);

        sheet.GetRange(totalRow, 0, totalRow, 2).Style = new CellStyle { Font = new Font { Bold = true } };
        sheet.GetRange(1, 2, totalRow + 1, 2).Style = new CellStyle { NumberFormat = "#,##0.00" };
        Widen(sheet);
        wb.Recalculate();
        return wb;
    }

    private static List<Trip> Trips() =>
    [
        new Trip("Lisbon", 4, 820),
        new Trip("Tallinn", 2, 410),
        new Trip("Osaka", 7, 2140),
        new Trip("Reykjavik", 3, 1290),
    ];

    private static Dictionary<string, object> Row(string key, int count, string state) => new()
    {
        ["key"] = key,
        ["count"] = count,
        ["state"] = state,
    };

    private static void Widen(Sheet sheet)
    {
        sheet.Columns[0].Width = 18;
        sheet.Columns[1].Width = 12;
        sheet.Columns[2].Width = 14;
    }

    /// <summary>
    /// A plain class: properties are discovered from the public getters (in name order, as the
    /// Java SDK orders JavaBean getters). A record would import in declaration order instead.
    /// </summary>
    public sealed class Trip(string city, int nights, double cost)
    {
        public string City { get; } = city;

        public int Nights { get; } = nights;

        public double Cost { get; } = cost;
    }
}
