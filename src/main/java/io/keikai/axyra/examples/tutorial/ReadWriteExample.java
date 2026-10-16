package io.keikai.axyra.examples.tutorial;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.FormulaPolicy;
import io.keikai.axyra.sheets.io.SaveOptions;

import java.nio.file.Path;

public final class ReadWriteExample implements Example {
    @Override
    public String name() {
        return "tutorial/read-write";
    }

    @Override
    public String description() {
        return "Open, edit, recalculate, and save an existing workbook";
    }

    @Override
    public void run(Path outputDirectory) {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("report.xlsx"));
        Path edited = outputDirectory.resolve("report-edited.xlsx");

        try (Workbook workbook = Workbook.open(source)) {
            for (int i = 0; i < workbook.sheetCount(); i++) {
                System.out.println("Sheet: " + workbook.sheet(i).name());
            }

            Sheet sheet = workbook.sheet("Q3 Sales");   // by name, not index
            sheet.cell(2, 2).setValue(CellValue.number(1499));
            Range q1 = sheet.range("B2:B4");
            q1.setNumbers(new double[] {1200, 980, 1610});
            sheet.setColumnWidth(0, 24);
            sheet.freezePanes(1, 0);

            // Edits recalculate their dependents; D2:D4 are already current.
            System.out.println("D3 after edit: " + sheet.value(2, 3));
            CellValue answer = workbook.evaluate("SUM('Q3 Sales'!B2:B4)");
            System.out.println("Q1 total: " + answer);

            workbook.save(edited, SaveOptions.create()
                    .formulaPolicy(FormulaPolicy.KEEP)
                    .skipComments(false)
                    .skipHyperlinks(false));
        }
        System.out.println("Wrote " + edited);
    }
}

