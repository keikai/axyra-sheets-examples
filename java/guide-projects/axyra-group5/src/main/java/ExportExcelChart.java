import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;
import java.nio.file.Files;
import java.nio.file.Path;

/** Exports an existing chart as PNG and its report sheet as PDF. */
public final class ExportExcelChart {
    public static void main(String[] args) throws Exception {
        try (Workbook wb = Workbook.open(Path.of("sales-chart.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            Files.write(Path.of("sales-chart.png"),
                    sales.renderChartToPng(0, 1200, 700));
            // Keep the smaller PDF chart readable. The XLSX remains unchanged.
            sales.chart(0).setShowDataLabels(false);
            sales.setPrintArea("A1:N19");
            sales.setLandscape(true);
            sales.fitToPage(1, 1);
            wb.renderPdf(Path.of("sales-chart.pdf"), PdfOptions.builder()
                    .sheetNames("Sales").build());
        }
        System.out.println("Wrote sales-chart.png and sales-chart.pdf");
    }
}
