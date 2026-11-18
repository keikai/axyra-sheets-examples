import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.ImageOptions;
import java.nio.file.*;

/** Renders actual saved workbooks; evaluation notices are retained. */
public final class RenderGuideFigures {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("expected"));
        render("input/sales-template.xlsx", "Invoice", "template.png");
        render("output/acme-sales.xlsx", "Invoice", "acme-sales.png");
    }
    private static void render(String input, String sheet, String output) throws Exception {
        try (Workbook wb = Workbook.open(Path.of(input))) {
            // Landscape preview keeps all five columns visible; the XLSX is not resaved.
            wb.sheet(sheet).setLandscape(true);
            Files.write(Path.of("expected", output), wb.renderSheetPng(
                    wb.sheet(sheet).index(), ImageOptions.builder().scale(1.5).gridLines(true).build()));
        }
    }
}
