import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;
import java.nio.file.Files;
import java.nio.file.Path;

/** Exports a chart image, a chart-only PDF, and a complete report PDF. */
public final class ExportChartFormats {
    public static void main(String[] args) throws Exception {
        Path input = Path.of(args.length > 0 ? args[0] : "sales-chart.xlsx");
        Path output = Path.of(args.length > 1 ? args[1] : "output/chart-export");
        Files.createDirectories(output);

        try (Workbook workbook = Workbook.open(input)) {
            Sheet sales = workbook.sheet("Sales");
            if (sales.charts().isEmpty()) {
                throw new IllegalArgumentException("Sales must contain a chart");
            }
            Files.write(output.resolve("sales-chart.png"),
                    sales.renderChartToPng(0, 1200, 700));

            sales.setLandscape(true);
            sales.fitToPage(1, 1);
            PdfOptions pdf = PdfOptions.builder().sheetNames("Sales").build();

            // Include the table and chart in the report page.
            sales.chart(0).setShowDataLabels(false);
            sales.setPrintArea("A1:N19");
            workbook.renderPdf(output.resolve("sales-report.pdf"), pdf);

            // Print the chart's anchored area while retaining its source cells.
            sales.setPrintArea("E2:N19");
            workbook.renderPdf(output.resolve("sales-chart-only.pdf"), pdf);
        }
        System.out.println("Wrote chart PNG and two PDF layouts to " + output);
    }
}
