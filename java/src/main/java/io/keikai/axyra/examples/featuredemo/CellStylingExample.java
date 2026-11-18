package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;
import io.keikai.axyra.sheets.style.Side;
import io.keikai.axyra.sheets.style.Underline;
import io.keikai.axyra.sheets.style.VAlign;

import java.nio.file.Path;

/**
 * Cell styling — fonts, fills, borders and alignment.
 *
 * <p>{@link CellStyle} is an immutable value built once and applied to a whole
 * range in one call. Identical styles are interned by the engine, so styling a
 * thousand rows with the same builder result costs one style entry in the file,
 * not a thousand.
 */
public final class CellStylingExample implements Example {
    private static final String CONTENT = "cellStyling-grid-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/cell-styling";
    }

    @Override
    public String description() {
        return "Style cells with fonts, fills, borders and alignment";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = fonts()) {
            wb.save(outputDirectory.resolve("cell-styling-fonts.xlsx"));
        }
        try (Workbook wb = fills()) {
            wb.save(outputDirectory.resolve("cell-styling-fills.xlsx"));
        }
        try (Workbook wb = borders()) {
            wb.save(outputDirectory.resolve("cell-styling-borders.xlsx"));
        }
        try (Workbook wb = alignment()) {
            wb.save(outputDirectory.resolve("cell-styling-alignment.xlsx"));
        }
        try (Workbook wb = styledReport()) {
            wb.save(outputDirectory.resolve("cell-styling-report.xlsx"));
        }
        System.out.println("Wrote five styling workbooks to " + outputDirectory);
    }

    private static Workbook fonts() {
        Workbook wb = grid("Fonts");
        Sheet sheet = wb.sheet(0);

        label(sheet, 1, "fontName(\"Georgia\")");
        sheet.range("B2").setStyle(CellStyle.builder().fontName("Georgia").fontSize(13).build());

        label(sheet, 2, "fontSize(18)");
        sheet.range("B3").setStyle(CellStyle.builder().fontSize(18).build());

        label(sheet, 3, "bold().italic()");
        sheet.range("B4").setStyle(CellStyle.builder().bold().italic().build());

        label(sheet, 4, "underline(DOUBLE)");
        sheet.range("B5").setStyle(CellStyle.builder().underline(Underline.DOUBLE).build());

        label(sheet, 5, "fontColor(dark red)");
        sheet.range("B6").setStyle(CellStyle.builder()
                .fontColor(Color.rgb(0xC0, 0x00, 0x00)).build());

        label(sheet, 6, "strikethrough(true)");
        sheet.range("B7").setStyle(CellStyle.builder().strikethrough(true).build());
        return wb;
    }

    private static Workbook fills() {
        Workbook wb = grid("Fills");
        Sheet sheet = wb.sheet(0);

        label(sheet, 1, "fill(solid)");
        sheet.range("B2").setStyle(CellStyle.builder()
                .fill(Color.rgb(0xDD, 0xEB, 0xF7)).build());

        label(sheet, 2, "fill(solid), any colour");
        sheet.range("B3").setStyle(CellStyle.builder()
                .fill(Color.rgb(0xFF, 0xF2, 0xCC)).build());

        label(sheet, 3, "fill(solid) + fontColor");
        sheet.range("B4").setStyle(CellStyle.builder()
                .fill(Color.rgb(0x1F, 0x4E, 0x79)).fontColor(Color.WHITE).build());

        sheet.setColumnWidth(0, 34);
        return wb;
    }

    private static Workbook borders() {
        Workbook wb = grid("Borders");
        Sheet sheet = wb.sheet(0);

        label(sheet, 1, "per edge, per style");
        sheet.range("B2").setStyle(CellStyle.builder()
                .border(Side.TOP, BorderStyle.MEDIUM, Color.BLACK)
                .border(Side.BOTTOM, BorderStyle.DOUBLE, Color.rgb(0xC0, 0x00, 0x00))
                .border(Side.LEFT, BorderStyle.THIN, Color.GRAY)
                .border(Side.RIGHT, BorderStyle.DASHED, Color.GRAY)
                .build());

        label(sheet, 3, "allBorders over a block ->");
        sheet.range("D4:G9").setStyle(CellStyle.builder()
                .allBorders(BorderStyle.THIN, Color.rgb(0x1F, 0x4E, 0x79)).build());

        sheet.setColumnWidth(0, 34);
        return wb;
    }

    private static Workbook alignment() {
        Workbook wb = grid("Alignment");
        Sheet sheet = wb.sheet(0);
        sheet.setRowHeight(1, 36);
        sheet.setRowHeight(2, 36);
        sheet.setRowHeight(3, 46);

        label(sheet, 1, "CENTER / CENTER");
        sheet.range("B2").setStyle(CellStyle.builder()
                .horizontalAlign(HAlign.CENTER).verticalAlign(VAlign.CENTER).build());

        label(sheet, 2, "RIGHT / BOTTOM");
        sheet.range("B3").setStyle(CellStyle.builder()
                .horizontalAlign(HAlign.RIGHT).verticalAlign(VAlign.BOTTOM).build());

        sheet.cell(3, 0).setValue(CellValue.text("wrapText + indent"));
        sheet.cell(3, 1).setValue(CellValue.text(
                "A longer sentence that has to wrap inside its own cell."));
        sheet.range("B4").setStyle(CellStyle.builder()
                .wrapText(true).verticalAlign(VAlign.TOP).indent(1).build());
        return wb;
    }

    /** The everyday case: one header style, one body style, applied by range. */
    private static Workbook styledReport() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Report");

        CellStyle header = CellStyle.builder()
                .bold().fontColor(Color.WHITE)
                .fill(Color.rgb(0x1F, 0x4E, 0x79))
                .horizontalAlign(HAlign.CENTER)
                .allBorders(BorderStyle.THIN, Color.rgb(0x1F, 0x4E, 0x79))
                .build();
        CellStyle money = CellStyle.builder()
                .numberFormat("#,##0.00")
                .allBorders(BorderStyle.HAIR, Color.rgb(0xBF, 0xBF, 0xBF))
                .build();

        sheet.range("A1:D1").setValues(new CellValue[][] {{
                CellValue.text("Client"), CellValue.text("Hours"),
                CellValue.text("Rate"), CellValue.text("Fee")}});
        String[] clients = {"Adler", "Brandt", "Cortez", "Duval", "Eriksen"};
        double[] hours = {42, 17.5, 63, 8, 29.25};
        double[] rate = {145, 190, 120, 210, 160};
        for (int i = 0; i < clients.length; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(clients[i]));
            sheet.setNumber(i + 1, 1, hours[i]);
            sheet.setNumber(i + 1, 2, rate[i]);
            sheet.cell(i + 1, 3).setFormula("B" + (i + 2) + "*C" + (i + 2));
        }

        sheet.range("A1:D1").setStyle(header);
        sheet.range(1, 1, clients.length, 3).setStyle(money);
        sheet.setColumnWidth(0, 16);
        sheet.setColumnWidth(3, 14);
        wb.recalculate();
        return wb;
    }

    /** The starting sheet: a labelled grid read from a packaged workbook, renamed for the sample. */
    private static Workbook grid(String name) {
        Workbook wb = Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
        wb.renameSheet(0, name);
        return wb;
    }

    private static void label(Sheet sheet, int row, String call) {
        sheet.cell(row, 0).setValue(CellValue.text(call));
        sheet.cell(row, 1).setValue(CellValue.text("Sample text 1234"));
    }
}
