import io.keikai.axyra.sheets.*;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.calc.CalcMode;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.*;
import java.util.Map;

/** Generates article evidence from real Axyra workbooks; no Excel UI is implied. */
public final class RenderReviewFigures {
    public static void main(String[] args) throws Exception {
        Path out = Path.of("output/review-figures");
        Files.createDirectories(out);
        try (Workbook wb = Workbook.open(Path.of("expected/sales-chart.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            Files.write(out.resolve("20-column-chart.png"), sales.renderChartToPng(0, 1200, 700));
            sales.addChart(Chart.pie().title("Revenue Share by Region")
                    .series("Actual", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5")
                    .legend(LegendPosition.RIGHT).showDataLabels(true)
                    .anchor(20, 4, 37, 13).build());
            Sheet monthly = wb.createSheet("Monthly");
            String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
            monthly.cell(0, 0).setValue(CellValue.text("Month"));
            monthly.cell(0, 1).setValue(CellValue.text("Revenue (USD)"));
            for (int i = 0; i < months.length; i++) {
                monthly.cell(i + 1, 0).setValue(CellValue.text(months[i]));
                monthly.cell(i + 1, 1).setValue(CellValue.number(1000 * (i + 1)));
            }
            monthly.addChart(Chart.line().title("Monthly Revenue")
                    .series("Revenue", "Monthly!$B$2:$B$13", "Monthly!$A$2:$A$13")
                    .categoryAxisTitle("Month").valueAxisTitle("Revenue (USD)")
                    .legendVisible(false).anchor(1, 3, 18, 12).build());
            Path variants = out.resolve("chart-variants.xlsx");
            wb.save(variants);
        }
        try (Workbook wb = Workbook.open(out.resolve("chart-variants.xlsx"))) {
            Files.write(out.resolve("20-pie-chart.png"), wb.sheet("Sales").renderChartToPng(1, 1200, 700));
            Files.write(out.resolve("20-line-chart.png"), wb.sheet("Monthly").renderChartToPng(0, 1200, 700));
        }
        try (Workbook wb = Workbook.open(Path.of("expected/formatted-sales.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            ImageOptions options = ImageOptions.builder().scale(1.5).build();
            Files.write(out.resolve("21-formatted-page.png"),
                    wb.renderRangePng(sales.index(), "A1:D5", options));
            try (Workbook raw = Workbook.create()) {
                Sheet plain = raw.createSheet("Sales");
                plain.range("A1:D5").setValues(sales.range("A1:D5").values());
                double[] widths = {16, 16, 22, 14};
                for (int c = 0; c < 4; c++) plain.setColumnWidth(c, widths[c]);
                Path before = out.resolve("unformatted-sales.xlsx");
                raw.save(before);
                Files.write(out.resolve("21-unformatted-page.png"),
                        raw.renderRangePng(plain.index(), "A1:D5", options));
            }
            new ObjectMapper().writeValue(out.resolve("format-evidence.json").toFile(), Map.of(
                    "revenueRaw", ((CellValue.Number) sales.cell(1, 2).value()).value(),
                    "revenueDisplay", sales.cell(1, 2).formattedText(),
                    "marginRaw", ((CellValue.Number) sales.cell(1, 3).value()).value(),
                    "marginDisplay", sales.cell(1, 3).formattedText(),
                    "dateRaw", ((CellValue.Number) sales.cell(1, 0).value()).value(),
                    "dateDisplay", sales.cell(1, 0).formattedText()));
        }
        try (Workbook wb = Workbook.open(Path.of("expected/sales-calculated.xlsx"))) {
            wb.setCalcMode(CalcMode.AUTOMATIC);
            Sheet sales = wb.sheet("Sales");
            sales.cell(1, 2).setValue(CellValue.number(2));
            double before = ((CellValue.Number) wb.sheet("Summary").cell(1, 8).value()).value();
            wb.save(out.resolve("formula-before.xlsx"));
            sales.cell(1, 2).setValue(CellValue.number(4));
            double after = ((CellValue.Number) wb.sheet("Summary").cell(1, 8).value()).value();
            if (before != 445 || after != 495) throw new AssertionError("Wrong formula preview totals");
            new ObjectMapper().writeValue(out.resolve("formula-evidence.json").toFile(),
                    Map.of("beforeQuantity", 2, "afterQuantity", 4, "unitPrice", 25,
                            "beforeTotal", before, "afterTotal", after));
        }
        System.out.println("Rendered native chart and workbook figures; verified formula and formatting evidence");
    }
}
