import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ChartView;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;

/** Adds a fifth region to the report created by CreateExcelCharts. */
public final class UpdateExcelChart {
    public static void main(String[] args) {
        Path output = Path.of("sales-chart-updated.xlsx");
        try (Workbook wb = Workbook.open(Path.of("sales-chart.xlsx"))) {
            Sheet sheet = wb.sheet("Sales");
            sheet.range("A6:C6").setValues(new CellValue[][] {
                {CellValue.text("Central"), CellValue.number(119000),
                    CellValue.number(118000)}
            });
            sheet.range("B6:C6").setStyle(sheet.cell(1, 1).style());
            ChartView chart = sheet.chart(0);
            chart.setTitle("Updated Actual vs Target Revenue");
            chart.setSeries(0, "Actual", "Sales!$B$2:$B$6", "Sales!$A$2:$A$6");
            chart.setSeries(1, "Target", "Sales!$C$2:$C$6", "Sales!$A$2:$A$6");
            wb.save(output);
        }
        try (Workbook check = Workbook.open(output)) {
            ChartView chart = check.sheet("Sales").chart(0);
            if (!"Sales!$B$2:$B$6".equals(chart.seriesValues(0))
                    || !"Sales!$A$2:$A$6".equals(chart.seriesCategories(1))) {
                throw new IllegalStateException("Expected five-region chart references");
            }
        }
        System.out.println("Wrote " + output);
    }
}
