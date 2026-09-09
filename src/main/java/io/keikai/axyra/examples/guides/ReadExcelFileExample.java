package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetType;
import io.keikai.axyra.sheets.Workbook;

import java.nio.file.Path;

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
        try (Workbook workbook = Workbook.open(source)) {
            for (int i = 0; i < workbook.sheetCount(); i++) {
                Sheet sheet = workbook.sheet(i);
                if (sheet.sheetType() != SheetType.WORKSHEET) {
                    continue;
                }
                Range used = sheet.usedRange();
                if (used == null) {
                    continue;
                }
                System.out.println("== " + sheet.name());
                for (CellValue[] row : used.values()) {
                    StringBuilder line = new StringBuilder();
                    for (CellValue value : row) {
                        if (line.length() > 0) line.append('\t');
                        line.append(text(value));
                    }
                    System.out.println(line);
                }
            }
        }
    }

    private static String text(CellValue value) {
        if (value instanceof CellValue.Number n) return String.valueOf(n.value());
        if (value instanceof CellValue.Text t) return t.value();
        if (value instanceof CellValue.Bool b) return String.valueOf(b.value());
        if (value instanceof CellValue.Error e) return e.display();
        if (value instanceof CellValue.Array) return "[array]";
        return "";
    }
}

