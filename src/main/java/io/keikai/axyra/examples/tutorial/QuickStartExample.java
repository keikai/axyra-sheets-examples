package io.keikai.axyra.examples.tutorial;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;

import java.nio.file.Path;

public final class QuickStartExample implements Example {
    @Override
    public String name() {
        return "tutorial/quick-start";
    }

    @Override
    public String description() {
        return "Create a workbook, calculate a formula, and save hello.xlsx";
    }

    @Override
    public void run(Path outputDirectory) {
        Path output = outputDirectory.resolve("hello.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Sheet1");
            sheet.range("A1:D2").setValues(new CellValue[][] {
                    {CellValue.text("Item"), CellValue.text("Qty"),
                            CellValue.text("Price"), CellValue.text("Total")},
                    {CellValue.text("Widget"), CellValue.number(12),
                            CellValue.number(4.25), CellValue.blank()}
            });
            sheet.cell(1, 3).setFormula("B2*C2");
            workbook.recalculate();

            CellValue total = sheet.value(1, 3);
            if (total instanceof CellValue.Number number) {
                System.out.println("Calculated total: " + number.value());
            }

            Range header = sheet.range("A1:D1");
            System.out.println("Header columns: " + header.colCount());
            workbook.save(output);
        }
        System.out.println("Wrote " + output);
    }
}
