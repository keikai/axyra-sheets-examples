package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

/**
 * Manage worksheets — create, insert, rename, reorder, copy, delete.
 *
 * <p>Renaming rewrites every formula that referenced the old name, and deleting
 * follows Excel's own activation rules.
 */
public final class ManageWorksheetsExample implements Example {
    @Override
    public String name() {
        return "feature-demo/manage-worksheets";
    }

    @Override
    public String description() {
        return "Create, insert, rename, reorder, copy, delete and colour worksheets";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = createAndInsert()) {
            wb.save(outputDirectory.resolve("manage-worksheets-create-insert.xlsx"));
        }
        try (Workbook wb = renameAndReorder()) {
            wb.save(outputDirectory.resolve("manage-worksheets-rename-reorder.xlsx"));
        }
        try (Workbook wb = copySheet()) {
            wb.save(outputDirectory.resolve("manage-worksheets-copy.xlsx"));
        }
        try (Workbook wb = deleteSheet()) {
            wb.save(outputDirectory.resolve("manage-worksheets-delete.xlsx"));
        }
        try (Workbook wb = colourTabs()) {
            wb.save(outputDirectory.resolve("manage-worksheets-tab-colours.xlsx"));
        }
        System.out.println("Wrote five worksheet-management workbooks to " + outputDirectory);
    }

    /** Append with {@code createSheet}, or place one with {@code insertSheet}. */
    private static Workbook createAndInsert() {
        Workbook wb = Workbook.create();
        wb.createSheet("Data");
        wb.createSheet("Report");
        Sheet cover = wb.insertSheet(0, "Cover");
        label(cover, "Cover", "insertSheet(0, \"Cover\") put this tab first.");
        label(wb.sheet("Data"), "Data", "createSheet appends at the end.");
        label(wb.sheet("Report"), "Report", "Three sheets, in creation order.");
        wb.setActiveSheet(cover.index());
        return wb;
    }

    /** A rename rewrites the formulas that pointed at the old name. */
    private static Workbook renameAndReorder() {
        Workbook wb = Workbook.create();
        Sheet data = wb.createSheet("Data");
        Sheet report = wb.createSheet("Report");
        data.range("A1:A4").setNumbers(new double[] {10, 20, 30, 40});

        label(report, "Report", "This total references the other sheet.");
        report.cell(2, 0).setValue(CellValue.text("Total"));
        report.cell(2, 1).setFormula("SUM(Data!A1:A4)");

        // The formula above now reads SUM(SourceData!A1:A4).
        wb.renameSheet(data.index(), "SourceData");
        report.cell(4, 0).setValue(CellValue.text("Formula after the rename:"));
        report.cell(5, 0).setValue(CellValue.text(report.cell(2, 1).formula()));

        wb.moveSheet(report.index(), 0);
        wb.recalculate();
        wb.setActiveSheet(0);
        return wb;
    }

    /** A deep copy: values, styles and formulas are re-interned into the target. */
    private static Workbook copySheet() {
        Workbook wb = Workbook.create();
        Sheet source = wb.createSheet("Template");
        label(source, "Template", "Styled once, then copied.");
        source.range("A4:C4").setValues(new CellValue[][] {{
                CellValue.text("Item"), CellValue.text("Qty"), CellValue.text("Price")}});
        source.range("A4:C4").setStyle(CellStyle.builder().bold()
                .fill(Color.rgb(0xDD, 0xEB, 0xF7)).build());

        wb.copySheet(source, "January");
        wb.copySheet(source, "February");
        wb.copySheet(source, "March");
        return wb;
    }

    /** Delete by index; the surviving tabs shift down. */
    private static Workbook deleteSheet() {
        Workbook wb = Workbook.create();
        wb.createSheet("Keep");
        Sheet scratch = wb.createSheet("Scratch");
        wb.createSheet("Also keep");

        label(wb.sheet("Keep"), "Keep", "Three sheets were created.");
        wb.deleteSheet(scratch.index());
        label(wb.sheet("Also keep"), "Also keep",
                "\"Scratch\" was deleted — " + wb.sheetCount() + " sheets remain.");
        return wb;
    }

    /** Tab colour is a sheet property, not a style. */
    private static Workbook colourTabs() {
        Workbook wb = Workbook.create();
        String[] names = {"Q1", "Q2", "Q3", "Q4"};
        Color[] colors = {
                Color.rgb(0x1F, 0x4E, 0x79), Color.rgb(0x2E, 0x75, 0xB6),
                Color.rgb(0x63, 0xBE, 0x7B), Color.rgb(0xC0, 0x00, 0x00)};
        for (int i = 0; i < names.length; i++) {
            Sheet sheet = wb.createSheet(names[i]);
            sheet.setTabColor(colors[i]);
            label(sheet, names[i], "setTabColor tints the sheet tab.");
        }
        return wb;
    }

    private static void label(Sheet sheet, String heading, String note) {
        sheet.cell(0, 0).setValue(CellValue.text(heading));
        sheet.range("A1").setStyle(CellStyle.builder().bold().fontSize(14).build());
        sheet.cell(1, 0).setValue(CellValue.text(note));
        sheet.setColumnWidth(0, 22);
    }
}
