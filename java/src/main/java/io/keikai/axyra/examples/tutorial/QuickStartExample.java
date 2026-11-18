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
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Sheet1");

            sheet.cell(0, 0).setValue(CellValue.text("Item"));
            sheet.cell(0, 1).setValue(CellValue.text("Qty"));
            sheet.cell(0, 2).setValue(CellValue.text("Price"));

            sheet.cell(1, 0).setValue(CellValue.text("Widget"));
            sheet.setNumber(1, 1, 12);
            sheet.setNumber(1, 2, 4.25);

            sheet.cell(1, 3).setFormula("B2*C2");   // calculated right away

            wb.save(output);

            // Step 3: read a value back.
            CellValue total = sheet.value(1, 3);
            if (total instanceof CellValue.Number number) {
                System.out.println(number.value());       // 51.0
            }
            System.out.println(sheet.cell(1, 3).formattedText()); // 51

            // Step 4: A1 notation.
            Range header = sheet.range("A1:D1");    // same as sheet.range(0, 0, 0, 3)
            System.out.println("Header columns: " + header.colCount());
        }
        System.out.println("Wrote " + output);
    }
}
