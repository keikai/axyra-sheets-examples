import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.formula.Functions;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public final class ModernFormulaBatch {
    public static void main(String[] args) throws Exception {
        Path input = Path.of(args.length > 0 ? args[0] : "input/sales-input.xlsx");
        Path output = Path.of(args.length > 1 ? args[1] : "output/sales-calculated.xlsx");
        Files.createDirectories(output.toAbsolutePath().getParent());
        for (String name : List.of("XLOOKUP", "FILTER", "UNIQUE", "SORT")) {
            if (!Functions.contains(name)) {
                throw new IllegalStateException("Missing function: " + name);
            }
        }

        try (Workbook wb = Workbook.open(input)) {
            Sheet sales = wb.sheet("Sales");
            wb.setCalcMode(CalcMode.AUTOMATIC);
            sales.range("D2:D5").setFormula(
                    "XLOOKUP(B2,Products!$A$2:$A$4,Products!$B$2:$B$4,0)");
            sales.range("E2:E5").setFormula("C2*D2");

            // Each dynamic-array formula has one anchor and an empty spill area.
            Sheet summary = wb.createSheet("Summary");
            summary.cell(0, 0).setValue(CellValue.text("Regions"));
            summary.cell(1, 0).setFormula("SORT(UNIQUE(Sales!A2:A5))");
            summary.cell(0, 2).setValue(CellValue.text("North orders"));
            summary.cell(1, 2).setFormula(
                    "FILTER(Sales!A2:E5,Sales!A2:A5=\"North\",\"No orders\")");
            summary.cell(0, 8).setValue(CellValue.text("Revenue"));
            summary.cell(1, 8).setFormula("SUM(Sales!E2:E5)");

            requireNumber(summary.cell(1, 8).value(), 445, "Initial revenue");
            // A late correction demonstrates automatic dependent recalculation.
            sales.cell(1, 2).setValue(CellValue.number(4));
            requireNumber(sales.cell(1, 4).value(), 100, "Corrected order");
            requireNumber(summary.cell(1, 8).value(), 495, "Corrected revenue");

            // Recalculate once after the edits and before saving the report.
            wb.recalculate();
            summary.setColumnWidth(0, 16);
            summary.setColumnWidth(2, 18);
            summary.setColumnWidth(8, 18);
            wb.save(output);
        }

        try (Workbook check = Workbook.open(output)) {
            Sheet sales = check.sheet("Sales");
            Sheet summary = check.sheet("Summary");
            requireNumber(sales.cell(1, 4).value(), 100, "Saved order");
            requireNumber(summary.cell(1, 8).value(), 495, "Saved revenue");
            requireText(summary.cell(1, 0).value(), "East");
            requireText(summary.cell(2, 0).value(), "North");
            requireText(summary.cell(3, 0).value(), "South");
            requireNumber(summary.cell(1, 6).value(), 100, "First North order");
            requireNumber(summary.cell(2, 6).value(), 75, "Second North order");
            if (!"C2*D2".equals(sales.cell(1, 4).formula())) {
                throw new IllegalStateException("Order formula was not retained");
            }
            System.out.println("Revenue: " + summary.cell(1, 8).formattedText());
            System.out.println("Regions: East, North, South");
            System.out.println("North order amounts: 100, 75");
        }
        System.out.println("Wrote " + output);
    }

    private static void requireNumber(CellValue value, double expected, String label) {
        if (!(value instanceof CellValue.Number n)
                || Math.abs(n.value() - expected) > 0.000001) {
            throw new IllegalStateException(label + ": " + value);
        }
    }

    private static void requireText(CellValue value, String expected) {
        if (!(value instanceof CellValue.Text t) || !expected.equals(t.value())) {
            throw new IllegalStateException("Expected " + expected + ", got " + value);
        }
    }
}
