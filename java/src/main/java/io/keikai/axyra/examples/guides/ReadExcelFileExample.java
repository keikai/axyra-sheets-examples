package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetType;
import io.keikai.axyra.sheets.Workbook;

import java.nio.file.Path;

/** The complete program from "How to Read an Excel File in Java". */
public final class ReadExcelFileExample implements Example {
    @Override
    public String name() {
        return "guides/read-excel-file";
    }

    @Override
    public String description() {
        return "Read each worksheet's used range and print it as TSV";
    }

    @Override
    public void run(Path outputDirectory) {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("read-source.xlsx"));
        try (Workbook wb = Workbook.open(source)) {
            for (int i = 0; i < wb.sheetCount(); i++) {
                Sheet sheet = wb.sheet(i);
                if (sheet.sheetType() != SheetType.WORKSHEET) {
                    continue;
                }

                Range used = sheet.usedRange();
                if (used == null) {
                    System.out.println("== " + sheet.name() + " (empty)");
                    continue;
                }

                System.out.printf("== %s  rows %d-%d, cols %d-%d%n",
                        sheet.name(),
                        used.firstRow(), used.lastRow(),
                        used.firstColumn(), used.lastColumn());

                for (CellValue[] row : used.values()) {
                    StringBuilder line = new StringBuilder();
                    for (CellValue v : row) {
                        if (line.length() > 0) {
                            line.append('\t');
                        }
                        line.append(text(v));
                    }
                    System.out.println(line);
                }
            }
        }
    }

    private static String text(CellValue v) {
        if (v instanceof CellValue.Number n) return String.valueOf(n.value());
        if (v instanceof CellValue.Text t)   return t.value();
        if (v instanceof CellValue.Bool b)   return String.valueOf(b.value());
        if (v instanceof CellValue.Error e)  return e.display();
        if (v instanceof CellValue.Array a)  return "[array]";
        return "";
    }
}
