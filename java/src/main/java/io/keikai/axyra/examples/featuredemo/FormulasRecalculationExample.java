package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.formula.Functions;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;

/**
 * Formulas and recalculation.
 *
 * <p>Setting a formula does not evaluate it — {@code recalculate()} does, and
 * {@code recalculateDirty()} does only the part that changed. The engine also
 * evaluates ad hoc, with no cell involved, and supports the modern dynamic-array
 * functions that spill into neighbouring cells.
 */
public final class FormulasRecalculationExample implements Example {
    private static final String CONTENT = "engineFormulas-invoice-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/formulas-recalculation";
    }

    @Override
    public String description() {
        return "Calculate a sheet, evaluate ad hoc, spill dynamic arrays and use manual calc mode";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = calculate()) {
            wb.save(outputDirectory.resolve("formulas-calculate.xlsx"));
        }
        try (Workbook wb = beforeAndAfter()) {
            wb.save(outputDirectory.resolve("formulas-before-after.xlsx"));
        }
        try (Workbook wb = adHoc()) {
            wb.save(outputDirectory.resolve("formulas-ad-hoc.xlsx"));
        }
        try (Workbook wb = dynamicArrays()) {
            wb.save(outputDirectory.resolve("formulas-dynamic-arrays.xlsx"));
        }
        try (Workbook wb = manualMode()) {
            wb.save(outputDirectory.resolve("formulas-manual-mode.xlsx"));
        }
        System.out.println("Wrote five formula workbooks to " + outputDirectory);
    }

    /** An ordinary calculated sheet: per-row formulas plus aggregates. */
    private static Workbook calculate() {
        Workbook wb = invoice();
        Sheet sheet = wb.sheet(0);

        for (int row = 1; row <= 5; row++) {
            sheet.cell(row, 3).setFormula("B" + (row + 1) + "*C" + (row + 1));
        }
        sheet.cell(7, 0).setValue(CellValue.text("Subtotal"));
        sheet.cell(7, 3).setFormula("SUM(D2:D6)");
        sheet.cell(8, 0).setValue(CellValue.text("VAT 20%"));
        sheet.cell(8, 3).setFormula("D8*0.2");
        sheet.cell(9, 0).setValue(CellValue.text("Total"));
        sheet.cell(9, 3).setFormula("D8+D9");
        sheet.range("A8:D10").setStyle(CellStyle.builder().bold().build());
        sheet.range("D2:D10").setStyle(CellStyle.builder().numberFormat("#,##0.00").build());

        wb.recalculate();
        return wb;
    }

    /** The same book, showing that a formula is inert until recalculation. */
    private static Workbook beforeAndAfter() {
        Workbook wb = invoice();
        Sheet sheet = wb.sheet(0);
        for (int row = 1; row <= 5; row++) {
            sheet.cell(row, 3).setFormula("B" + (row + 1) + "*C" + (row + 1));
        }

        sheet.cell(7, 0).setValue(CellValue.text("D2 before recalculate()"));
        sheet.cell(7, 3).setValue(CellValue.text(String.valueOf(sheet.cell(1, 3).value())));

        wb.recalculate();

        sheet.cell(8, 0).setValue(CellValue.text("D2 after recalculate()"));
        sheet.cell(8, 3).setValue(CellValue.text(String.valueOf(sheet.cell(1, 3).value())));
        sheet.setColumnWidth(0, 26);
        sheet.setColumnWidth(3, 26);
        return wb;
    }

    /** {@code evaluate} needs no cell — useful for validating user input. */
    private static Workbook adHoc() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Evaluate");
        sheet.range("A1:A5").setNumbers(new double[] {10, 20, 30, 40, 50});
        sheet.cell(0, 1).setValue(CellValue.text("<- source data"));

        String[] expressions = {
                "1+2*3",
                "TEXT(0.1234, \"0.0%\")",
                "SUM(A1:A5)",
                "AVERAGE(A1:A5)",
                "XLOOKUP(30, A1:A5, A1:A5)",
                "LET(n, SUM(A1:A5), n/COUNT(A1:A5))"};

        sheet.range("C1:D1").setValues(new CellValue[][] {{
                CellValue.text("Expression"), CellValue.text("evaluate(...)")}});
        sheet.range("C1:D1").setStyle(CellStyle.builder().bold().build());
        for (int i = 0; i < expressions.length; i++) {
            sheet.cell(i + 1, 2).setValue(CellValue.text(expressions[i]));
            sheet.cell(i + 1, 3).setValue(CellValue.text(
                    String.valueOf(wb.evaluate(expressions[i], sheet))));
        }

        sheet.cell(9, 2).setValue(CellValue.text("Functions in the catalogue"));
        sheet.setNumber(9, 3, Functions.count());
        sheet.setColumnWidth(2, 38);
        sheet.setColumnWidth(3, 30);
        return wb;
    }

    /** One formula, many cells: the result spills into the neighbours. */
    private static Workbook dynamicArrays() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Spill");

        sheet.range("A1:B1").setValues(new CellValue[][] {{
                CellValue.text("Region"), CellValue.text("Revenue")}});
        String[] regions = {"North", "South", "East", "West", "North", "East"};
        double[] revenue = {1800, 1470, 2415, 1695, 2200, 1310};
        for (int i = 0; i < regions.length; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(regions[i]));
            sheet.setNumber(i + 1, 1, revenue[i]);
        }
        sheet.range("A1:B1").setStyle(CellStyle.builder().bold().build());

        sheet.cell(0, 3).setValue(CellValue.text("SEQUENCE(6)"));
        sheet.cell(1, 3).setFormula("SEQUENCE(6)");

        sheet.cell(0, 4).setValue(CellValue.text("UNIQUE(A2:A7)"));
        sheet.cell(1, 4).setFormula("UNIQUE(A2:A7)");

        sheet.cell(0, 5).setValue(CellValue.text("SORT(B2:B7,1,-1)"));
        sheet.cell(1, 5).setFormula("SORT(B2:B7,1,-1)");

        sheet.cell(0, 6).setValue(CellValue.text("FILTER(A2:A7,B2:B7>1700)"));
        sheet.cell(1, 6).setFormula("FILTER(A2:A7,B2:B7>1700)");

        sheet.range("D1:G1").setStyle(CellStyle.builder().bold().build());
        for (int c = 3; c <= 6; c++) {
            sheet.setColumnWidth(c, 20);
        }
        wb.recalculate();
        return wb;
    }

    /** Manual mode defers evaluation until you ask, like Excel's F9. */
    private static Workbook manualMode() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Manual");
        wb.setCalcMode(CalcMode.MANUAL);

        sheet.range("A1:A5").setNumbers(new double[] {3, 1, 4, 1, 5});
        sheet.cell(6, 0).setValue(CellValue.text("Calc mode"));
        sheet.cell(6, 1).setValue(CellValue.text(String.valueOf(wb.calcMode())));
        sheet.cell(7, 0).setValue(CellValue.text("SUM(A1:A5)"));
        sheet.cell(7, 1).setFormula("SUM(A1:A5)");

        sheet.cell(9, 0).setValue(CellValue.text(
                "In MANUAL mode the total stays stale until recalculate() runs."));
        wb.recalculate();
        sheet.cell(10, 0).setValue(CellValue.text("recalculate() has now run."));
        sheet.setColumnWidth(0, 30);
        return wb;
    }

    /** The starting sheet: invoice lines read from a packaged workbook. */
    private static Workbook invoice() {
        return Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
    }
}
