package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Save and export — one workbook, several formats, no temporary files.
 *
 * <p>{@code saveBytes(format)} serializes straight to a {@code byte[]}, which is
 * what a web application actually wants: the same call feeds a download, a blob
 * column or an HTTP response. {@code save(Path)} picks the writer from the file
 * extension instead.
 */
public final class SaveExportExample implements Example {
    private static final String[] FORMATS = {"xlsx", "csv", "json", "html"};

    @Override
    public String name() {
        return "feature-demo/save-export";
    }

    @Override
    public String description() {
        return "Serialize one workbook to xlsx, csv, json and html bytes";
    }

    @Override
    public void run(Path outputDirectory) throws IOException {
        for (String format : FORMATS) {
            byte[] bytes = export(format);
            Path file = outputDirectory.resolve("quarterly-report." + format);
            Files.write(file, bytes);
            System.out.println("Wrote " + file.getFileName() + " (" + bytes.length + " bytes)");
        }
    }

    /** Serialize to {@code format}, ready for a download, an attachment or a blob column. */
    private static byte[] export(String format) {
        try (Workbook wb = report()) {
            return wb.saveBytes(format);
        }
    }

    /** A small report worth exporting: labels, numbers and a live total. */
    private static Workbook report() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Q3");

        sheet.range("A1:C1").setValues(new CellValue[][] {{
                CellValue.text("Region"), CellValue.text("Units"), CellValue.text("Revenue")}});
        sheet.range("A1:C1").setStyle(CellStyle.builder().bold().build());

        String[] regions = {"North", "South", "East", "West"};
        double[] units = {1240, 980, 1610, 1130};
        double[] revenue = {18600, 14700, 24150, 16950};
        for (int i = 0; i < regions.length; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(regions[i]));
            sheet.setNumber(i + 1, 1, units[i]);
            sheet.setNumber(i + 1, 2, revenue[i]);
        }

        sheet.cell(5, 0).setValue(CellValue.text("Total"));
        sheet.cell(5, 1).setFormula("SUM(B2:B5)");
        sheet.cell(5, 2).setFormula("SUM(C2:C5)");
        sheet.range("A6:C6").setStyle(CellStyle.builder().bold().build());
        sheet.range("C2:C6").setStyle(CellStyle.builder().numberFormat("#,##0").build());
        sheet.setColumnWidth(0, 14);
        sheet.setColumnWidth(2, 14);

        // CSV, JSON and HTML all read the calculated values, so evaluate first.
        wb.recalculate();
        return wb;
    }
}
