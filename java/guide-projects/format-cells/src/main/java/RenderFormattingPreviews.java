import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.BorderStyle;
import java.nio.file.Files;
import java.nio.file.Path;

/** Generates previews for the format-cells examples. */
public final class RenderFormattingPreviews {
    public static void main(String[] args) throws Exception {
        Path out = Path.of("output/previews");
        Files.createDirectories(out);
        ImageOptions image = ImageOptions.builder().scale(1.5).build();
        try (Workbook wb = Workbook.open(Path.of("formatted-sales.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            sales.setLandscape(true);
            Files.write(out.resolve("formatted-report.png"),
                    wb.renderRangePng(sales.index(), "A1:D5", image));
            try (Workbook raw = Workbook.create()) {
                Sheet plain = raw.createSheet("Sales");
                plain.range("A1:D5").setValues(sales.range("A1:D5").values());
                // A neutral grid and readable dates keep the raw-data preview clear.
                plain.range("A1:D5").setStyle(CellStyle.builder()
                        .fontName("Calibri").fontSize(11)
                        .allBorders(BorderStyle.THIN, Color.parse("#D9D9D9"))
                        .build());
                plain.range("A2:A4").setStyle(plain.cell(1, 0).style().toBuilder()
                        .numberFormat("yyyy-mm-dd").build());
                for (int row = 0; row < 5; row++) plain.setRowHeight(row, 22);
                double[] widths = {16, 16, 22, 14};
                for (int c = 0; c < 4; c++) plain.setColumnWidth(c, widths[c]);
                plain.setLandscape(true);
                raw.save(out.resolve("unformatted-sales.xlsx"));
                Files.write(out.resolve("unformatted-report.png"),
                        raw.renderRangePng(plain.index(), "A1:D5", image));
            }
        }
        try (Workbook wb = Workbook.open(Path.of("formatted-sales-highlighted.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            sales.setLandscape(true);
            Files.write(out.resolve("highlighted-report.png"),
                    wb.renderRangePng(sales.index(), "A1:D5", image));
        }
        System.out.println("Wrote previews to " + out);
    }
}
