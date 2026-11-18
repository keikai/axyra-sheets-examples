package io.keikai.axyra.examples;

import io.keikai.axyra.examples.tutorial.QuickStartExample;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class QuickStartIntegrationTest {
    @Test
    void writesARealWorkbook(@TempDir Path output) throws Exception {
        new QuickStartExample().run(output);
        Path workbookPath = output.resolve("hello.xlsx");
        assertTrue(Files.size(workbookPath) > 0);
        try (Workbook workbook = Workbook.open(workbookPath)) {
            assertEquals("Widget", ((CellValue.Text) workbook.sheet("Sheet1")
                    .value(1, 0)).value());
            assertEquals(51.0, ((CellValue.Number) workbook.sheet("Sheet1")
                    .value(1, 3)).value());
        }
    }
}
