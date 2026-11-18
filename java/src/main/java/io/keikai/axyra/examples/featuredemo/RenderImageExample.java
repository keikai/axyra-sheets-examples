package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.io.Renderer;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Render a sheet to an image.
 *
 * <p>{@link Renderer} rasterises a sheet server-side — PNG, JPEG, GIF, BMP, TIFF
 * or SVG — with no browser and no viewer component involved. That is what makes
 * it useful for thumbnails, email bodies and report previews.
 */
public final class RenderImageExample implements Example {
    @Override
    public String name() {
        return "feature-demo/render-image";
    }

    @Override
    public String description() {
        return "Render a sheet to PNG at different scales, with or without gridlines, and to SVG";
    }

    @Override
    public void run(Path outputDirectory) throws IOException {
        try (Workbook wb = dashboard()) {
            wb.save(outputDirectory.resolve("render-image-dashboard.xlsx"));
        }
        write(outputDirectory.resolve("sheet.png"), render("png", 1.0, true));
        write(outputDirectory.resolve("sheet-2x.png"), render("png", 2.0, true));
        write(outputDirectory.resolve("sheet-no-gridlines.png"), render("png", 1.5, false));
        write(outputDirectory.resolve("sheet.svg"), renderSvg());
    }

    private static void write(Path file, byte[] bytes) throws IOException {
        Files.write(file, bytes);
        System.out.println("Wrote " + file.getFileName() + " (" + bytes.length + " bytes)");
    }

    /** Rasterise the first sheet at the given scale. */
    private static byte[] render(String format, double scale, boolean gridLines) {
        try (Workbook wb = dashboard()) {
            return Renderer.toImage(wb, 0, format, ImageOptions.builder()
                    .scale(scale)
                    .gridLines(gridLines)
                    .build());
        }
    }

    /** SVG stays vector. */
    private static byte[] renderSvg() {
        try (Workbook wb = dashboard()) {
            return Renderer.toSvg(wb, 0, ImageOptions.builder().build());
        }
    }

    /** Something worth looking at: figures, a total and a chart. */
    private static Workbook dashboard() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Dashboard");

        sheet.range("A1:C1").setValues(new CellValue[][] {{
                CellValue.text("Month"), CellValue.text("Plan"), CellValue.text("Actual")}});
        sheet.range("A1:C1").setStyle(CellStyle.builder()
                .bold().fontColor(Color.WHITE)
                .fill(Color.rgb(0x1F, 0x4E, 0x79)).build());

        String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun"};
        double[] plan = {12000, 12500, 13000, 13500, 14000, 14500};
        double[] actual = {11400, 13200, 12750, 15100, 13800, 16250};
        for (int i = 0; i < months.length; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(months[i]));
            sheet.setNumber(i + 1, 1, plan[i]);
            sheet.setNumber(i + 1, 2, actual[i]);
        }
        sheet.range(1, 1, months.length, 2).setStyle(
                CellStyle.builder().numberFormat("#,##0").build());

        sheet.cell(8, 0).setValue(CellValue.text("Total"));
        sheet.cell(8, 1).setFormula("SUM(B2:B7)");
        sheet.cell(8, 2).setFormula("SUM(C2:C7)");
        sheet.range("A9:C9").setStyle(CellStyle.builder()
                .bold().numberFormat("#,##0").build());

        sheet.addChart(Chart.column()
                .title("Plan vs actual")
                .series("Plan", "Dashboard!B2:B7", "Dashboard!A2:A7")
                .series("Actual", "Dashboard!C2:C7", "Dashboard!A2:A7")
                .legend(LegendPosition.BOTTOM)
                .anchor(1, 4, 18, 11)
                .build());

        sheet.setColumnWidth(0, 12);
        sheet.setColumnWidth(1, 12);
        sheet.setColumnWidth(2, 12);

        // An image is one printed page, so lay the page out around the content:
        // the table and the chart beside it, landscape, scaled onto a single page.
        sheet.setPrintArea("A1:K19");
        sheet.setLandscape(true);
        sheet.fitToPage(1, 1);
        wb.recalculate();
        return wb;
    }
}
