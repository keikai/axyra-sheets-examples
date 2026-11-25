import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import java.nio.file.Files;
import java.nio.file.Path;

/** Generates previews for the create-charts examples. */
public final class RenderChartPreviews {
    public static void main(String[] args) throws Exception {
        Path out = Path.of("output/previews");
        Files.createDirectories(out);
        try (Workbook wb = Workbook.open(Path.of("sales-chart.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            Files.write(out.resolve("column-chart.png"),
                    sales.renderChartToPng(0, 1200, 700));
            sales.addChart(Chart.pie().title("Revenue Share by Region")
                    .series("Actual", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5")
                    .legend(LegendPosition.RIGHT).showDataLabels(true)
                    .anchor(20, 4, 37, 13).build());
            Sheet monthly = wb.createSheet("Monthly");
            String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
            double[] revenue = {9200, 8600, 10500, 11300, 10800, 12600,
                    11900, 11200, 13400, 14100, 16200, 17800};
            monthly.cell(0, 0).setValue(CellValue.text("Month"));
            monthly.cell(0, 1).setValue(CellValue.text("Revenue (USD)"));
            for (int i = 0; i < months.length; i++) {
                monthly.cell(i + 1, 0).setValue(CellValue.text(months[i]));
                monthly.cell(i + 1, 1).setValue(CellValue.number(revenue[i]));
            }
            monthly.addChart(Chart.line().title("Monthly Revenue")
                    .series("Revenue", "Monthly!$B$2:$B$13", "Monthly!$A$2:$A$13")
                    .categoryAxisTitle("Month").valueAxisTitle("Revenue (USD)")
                    .legendVisible(false).anchor(1, 3, 18, 12).build());
            wb.save(out.resolve("chart-variants.xlsx"));
        }
        try (Workbook wb = Workbook.open(out.resolve("chart-variants.xlsx"))) {
            Files.write(out.resolve("pie-chart.png"),
                    wb.sheet("Sales").renderChartToPng(1, 1200, 700));
            Files.write(out.resolve("line-chart.png"),
                    wb.sheet("Monthly").renderChartToPng(0, 1200, 700));
        }
        System.out.println("Wrote previews to " + out);
    }
}
