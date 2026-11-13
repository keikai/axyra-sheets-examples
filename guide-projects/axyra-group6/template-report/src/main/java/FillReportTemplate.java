import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.template.TemplateMarkers;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

public final class FillReportTemplate {
    public record Line(String sku, String product, int qty,
                       double unitPrice, double amount) {
        public static Line of(String sku, String product, int qty, double price) {
            return new Line(sku, product, qty, price, qty * price);
        }
    }

    public static void main(String[] args) throws Exception {
        Path input = Path.of("input/sales-template.xlsx");
        Path output = Path.of("output/acme-sales.xlsx");
        List<Line> lines = List.of(
                Line.of("00042", "Keyboard", 2, 49.5),
                Line.of("00043", "Mouse", 3, 25),
                Line.of("00044", "Monitor", 1, 220));
        Files.createDirectories(output.getParent());
        fill(input, output, lines);
        verify(output, lines);
        System.out.println("3 lines; total USD 394.00; wrote " + output);
    }

    static void fill(Path input, Path output, List<Line> lines) {
        if (lines.isEmpty()) {
            throw new IllegalArgumentException("A sales report needs at least one line");
        }
        try (Workbook wb = Workbook.open(input)) {
            Sheet sheet = wb.sheet("Invoice");
            Map<String, Object> data = Map.of(
                    "customer", "Acme Retail",
                    "period", "2026-10",
                    "lines", lines);
            TemplateMarkers.process(sheet, data);
            wb.recalculate();
            wb.save(output);
        }
    }

    static void verify(Path output, List<Line> lines) {
        try (Workbook wb = Workbook.open(output)) {
            Sheet sheet = wb.sheet("Invoice");
            double total = 0;
            for (int i = 0; i < lines.size(); i++) {
                int row = 4 + i;
                Line line = lines.get(i);
                if (!(sheet.cell(row, 0).value() instanceof CellValue.Text id)
                        || !line.sku().equals(id.value())) {
                    throw new IllegalStateException("Text SKU lost on row " + row);
                }
                requireNumber(sheet.cell(row, 4).value(), line.amount());
                String expected = "C" + (row + 1) + "*D" + (row + 1);
                if (!expected.equals(sheet.cell(row, 4).formula())) {
                    throw new IllegalStateException("Wrong detail formula");
                }
                if (!"#,##0.00".equals(sheet.cell(row, 4).style().numberFormat())) {
                    throw new IllegalStateException("Template number format lost");
                }
                total += line.amount();
            }
            requireNumber(sheet.cell(4 + lines.size(), 4).value(), total);
            for (CellValue[] row : sheet.usedRange().values()) {
                for (CellValue value : row) {
                    if (value instanceof CellValue.Text text && text.value().contains("${")) {
                        throw new IllegalStateException("Unresolved marker: " + text.value());
                    }
                }
            }
        }
    }

    private static void requireNumber(CellValue value, double expected) {
        if (!(value instanceof CellValue.Number n)
                || Math.abs(n.value() - expected) > 0.000001) {
            throw new IllegalStateException("Expected " + expected + ", got " + value);
        }
    }
}
