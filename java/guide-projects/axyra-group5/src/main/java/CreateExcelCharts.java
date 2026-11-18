import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

public final class CreateExcelCharts {
    public static void main(String[] args) {
        Path output = Path.of("sales-chart.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Sales");
            sheet.range("A1:C5").setValues(new CellValue[][] {
                {CellValue.text("Region"), CellValue.text("Actual (USD)"),
                    CellValue.text("Target (USD)")},
                {CellValue.text("North"), CellValue.number(125000),
                    CellValue.number(120000)},
                {CellValue.text("South"), CellValue.number(98000),
                    CellValue.number(105000)},
                {CellValue.text("East"), CellValue.number(142000),
                    CellValue.number(135000)},
                {CellValue.text("West"), CellValue.number(110000),
                    CellValue.number(115000)}
            });

            CellStyle header = CellStyle.builder()
                    .bold()
                    .fontColor(Color.parse("#FFFFFF"))
                    .fill(Color.parse("#1F3B68"))
                    .build();
            sheet.range("A1:C1").setStyle(header);
            sheet.range("B2:C5").setStyle(CellStyle.builder()
                    .numberFormat("#,##0").build());
            sheet.setColumnWidth(0, 16);
            sheet.setColumnWidth(1, 18);
            sheet.setColumnWidth(2, 18);
            sheet.freezePanes(1, 0);

            Chart chart = Chart.column()
                    .title("Actual vs Target Revenue")
                    .grouping("clustered")
                    .series("Actual", "Sales!$B$2:$B$5",
                            "Sales!$A$2:$A$5", Color.parse("#2F5597"))
                    .series("Target", "Sales!$C$2:$C$5",
                            "Sales!$A$2:$A$5", Color.parse("#ED7D31"))
                    .categoryAxisTitle("Region")
                    .valueAxisTitle("Revenue (USD)")
                    .legend(LegendPosition.BOTTOM)
                    .showDataLabels(true)
                    .chartAreaFill(Color.parse("#FFFFFF"))
                    .anchor(1, 4, 18, 13)
                    .build();

            sheet.addChart(chart);
            wb.save(output);
        }

        // Check the file that will be delivered.
        try (Workbook check = Workbook.open(output)) {
            Sheet sales = check.sheet("Sales");
            if (sales.charts().size() != 1
                    || sales.chart(0).seriesCount() != 2) {
                throw new IllegalStateException("Expected one two-series chart");
            }
        }
        System.out.println("Wrote " + output.toAbsolutePath());
    }
}
