package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.Cell;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SortKey;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

public final class CellAndRangeExample implements Example {
    record Sale(String region, double q1, double q2) {
    }

    @Override
    public String name() {
        return "dev-ref/cell-and-range";
    }

    @Override
    public String description() {
        return "Use typed cells, A1 ranges, bulk access, sorting, and record import";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.sheet(0);
            CellStyle bold = CellStyle.builder().bold().build();
            sheet.importData(0, 0, List.of(
                            new Sale("EMEA", 980, 1100),
                            new Sale("APAC", 1200, 1450)),
                    ImportOptions.builder()
                            .columns("region", "q1", "q2")
                            .headers(true)
                            .columnLabels("Region", "Q1", "Q2")
                            .headerStyle(bold)
                            .build());

            Range values = sheet.range("B2:C3");
            double[] numbers = values.numbers();
            values.setNumbers(numbers);
            sheet.range("A2:C3").sort(SortKey.asc(0));

            Cell date = sheet.cell(4, 0);
            date.setDate(LocalDate.of(2026, 3, 1));
            date.setStyle(CellStyle.builder().numberFormat("yyyy-mm-dd").build());
            sheet.cell(1, 3).setFormula("B2+C2");
            sheet.range("D2:D2").autoFill(sheet.range("D2:D3"));
            workbook.recalculate();

            CellValue calculated = sheet.value(1, 3);
            System.out.println(describe(calculated));
            workbook.save(outputDirectory.resolve("cell-and-range.xlsx"));
        }
    }

    private static String describe(CellValue value) {
        if (value instanceof CellValue.Number n) return "number " + n.value();
        if (value instanceof CellValue.Text t) return "text " + t.value();
        if (value instanceof CellValue.Bool b) return "bool " + b.value();
        if (value instanceof CellValue.Error e) return "error " + e.display();
        if (value instanceof CellValue.Array a) return a.rowCount() + "x" + a.colCount() + " array";
        return "blank";
    }
}

