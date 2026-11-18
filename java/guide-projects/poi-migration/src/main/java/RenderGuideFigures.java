import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.ImageOptions;
import java.nio.file.*;

/** Renders actual saved workbooks; evaluation notices are retained. */
public final class RenderGuideFigures {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("expected"));
        render("output/poi-sales.xlsx", "Sales", "poi-sales.png");
        render("output/axyra-sales.xlsx", "Sales", "axyra-sales.png");
    }
    private static void render(String input, String sheet, String output) throws Exception {
        try (Workbook wb = Workbook.open(Path.of(input))) {
            Files.write(Path.of("expected", output), wb.renderSheetPng(
                    wb.sheet(sheet).index(), ImageOptions.builder().scale(1.5).gridLines(true).build()));
        }
    }
}
