package io.keikai.axyra.examples;

import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public final class ExampleSupport {
    private ExampleSupport() {
    }

    public static Path prepareOutputDirectory(Path directory) throws IOException {
        return Files.createDirectories(directory.toAbsolutePath().normalize());
    }

    public static Path createSalesWorkbook(Path output) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Q3 Sales");
            sheet.range("A1:D4").setValues(new CellValue[][] {
                    {text("Region"), text("Q1"), text("Q2"), text("Total")},
                    {text("APAC"), number(1200), number(1450), CellValue.blank()},
                    {text("EMEA"), number(980), number(1100), CellValue.blank()},
                    {text("Americas"), number(1610), number(1720), CellValue.blank()}
            });
            sheet.range("D2:D4").setFormula("B2+C2");   // D3 = B3+C3, D4 = B4+C4
            workbook.save(output);
        }
        return output;
    }

    public static CellValue text(String value) {
        return CellValue.text(value);
    }

    public static CellValue number(double value) {
        return CellValue.number(value);
    }
}
