using System.Text.Json;
using System.Text.Json.Nodes;
using Axyra.Sheets.IO;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.Guides;

/// <summary>
/// The code from "How to Write Data to Excel in .NET". The guide maps a database query
/// to records; this example starts from the records so that it runs without a database.
/// </summary>
public sealed class WriteDataToExcelExample : IExample
{
    public sealed record Sale(string Region, string Product, DateOnly OrderDate, double Amount);

    private const string Json = """
        [{"region":"Americas","product":"Widget","amount":1720.0},
         {"region":"APAC","amount":null}]
        """;

    public string Name => "guides/write-data-to-excel";

    public string Description => "Import .NET records and JSON, and write a streaming Excel export";

    public void Run(string outputDirectory)
    {
        List<Sale> rows =
        [
            new Sale("APAC", "Widget", new DateOnly(2026, 9, 1), 1450),
            new Sale("EMEA", "Gadget", new DateOnly(2026, 9, 2), 1100),
        ];

        using (Workbook wb = new Workbook())
        {
            // Records: pinned columns, header row, then format the date and money columns.
            Sheet sheet = wb.Sheets.Add("Q3 Sales");
            var headerStyle = new CellStyle { Font = new Font { Bold = true } };
            int written = sheet.ImportData(0, 0, rows, new ImportOptions
            {
                Columns = ["Region", "Product", "OrderDate", "Amount"],
                Headers = true,
                ColumnLabels = ["Region", "Product", "Order Date", "Amount"],
                HeaderStyle = headerStyle,
            });
            Console.WriteLine(written + " rows including the header");

            var dateStyle = new CellStyle { NumberFormat = "yyyy-mm-dd" };
            var money = new CellStyle { NumberFormat = "#,##0.00" };
            int lastRow = rows.Count;                              // header is row 0
            sheet.GetRange(1, 2, lastRow, 2).Style = dateStyle;    // column C
            sheet.GetRange(1, 3, lastRow, 3).Style = money;        // column D

            // JSON parsed into dictionaries, imported directly. System.Text.Json reads object
            // values as JsonElement, so they are unwrapped to plain .NET values first.
            List<Dictionary<string, object?>> maps = JsonSerializer
                .Deserialize<List<Dictionary<string, JsonElement>>>(Json)!
                .Select(m => m.ToDictionary(e => e.Key, e => Plain(e.Value)))
                .ToList();
            Sheet fromMaps = wb.Sheets.Add("JSON (maps)");
            fromMaps.ImportData(0, 0, maps, new ImportOptions
            {
                Columns = ["region", "product", "amount"],
                ColumnLabels = ["Region", "Product", "Amount"],
            });

            // JSON with full control over each cell's type.
            Sheet dynamic = wb.Sheets.Add("JSON (dynamic)");
            JsonArray array = JsonNode.Parse(Json)!.AsArray();
            string[] columns = ["region", "product", "amount"];

            var header = new CellValue[1, columns.Length];
            for (int c = 0; c < columns.Length; c++)
            {
                header[0, c] = columns[c];
            }
            dynamic.GetRange(0, 0, 0, columns.Length - 1).SetValues(header);

            var body = new CellValue[array.Count, columns.Length];
            for (int r = 0; r < array.Count; r++)
            {
                JsonObject node = array[r]!.AsObject();
                for (int c = 0; c < columns.Length; c++)
                {
                    JsonNode? field = node[columns[c]];
                    body[r, c] = field?.GetValueKind() switch
                    {
                        null or JsonValueKind.Null => CellValue.Empty,
                        JsonValueKind.Number => field.GetValue<double>(),
                        JsonValueKind.True or JsonValueKind.False => field.GetValue<bool>(),
                        _ => field.ToString(),
                    };
                }
            }
            dynamic.GetRange(1, 0, array.Count, columns.Length - 1).SetValues(body);

            // The fast path for numbers: a flat, row-major span of doubles.
            int rowCount = 1_000, colCount = 12;
            double[] flat = new double[rowCount * colCount];
            for (int i = 0; i < flat.Length; i++)
            {
                flat[i] = i;                                       // flat[r * colCount + c]
            }
            Sheet matrix = wb.Sheets.Add("Numeric Matrix");
            matrix.GetRange(1, 0, rowCount, colCount - 1).SetNumbers(flat);

            wb.Save(Path.Combine(outputDirectory, "data-export.xlsx"));
        }

        WriteStreamingExport(Path.Combine(outputDirectory, "streaming-export.xlsx"), rows);
    }

    /// <summary>A JSON value as the plain .NET value a dictionary import understands.</summary>
    private static object? Plain(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Number => value.GetDouble(),
        JsonValueKind.String => value.GetString(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null or JsonValueKind.Undefined => null,
        _ => value.GetRawText(),
    };

    /// <summary>The streaming writer from the guide, fed from records instead of a query.</summary>
    private static void WriteStreamingExport(string output, List<Sale> sales)
    {
        using StreamWorkbook @out = new StreamWorkbook(output);
        int header = @out.AddStyle(new StreamStyle { Bold = true });
        int money = @out.AddStyle(new StreamStyle { NumberFormat = "#,##0.00" });
        int date = @out.AddStyle(new StreamStyle { NumberFormat = "yyyy-mm-dd" });

        @out.StartSheet("Sales");

        @out.StartRow(0).AddText("Region", header)
                        .AddText("Order Date", header)
                        .AddText("Amount", header)
                        .Commit();

        int r = 1;
        double total = 0;
        foreach (Sale sale in sales)
        {
            @out.StartRow(r++)
                .AddText(sale.Region)
                .AddDate(sale.OrderDate, date)
                .AddNumber(sale.Amount, money)
                .Commit();
            total += sale.Amount;
        }
        // No recalculation in a stream: supply the cached value with the formula.
        @out.StartRow(r).SetText(0, "Total", header)
                        .SetFormula(2, "SUM(C2:C" + r + ")", total, money)
                        .Commit();
    }
}
