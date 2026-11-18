package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.formula.Functions;

import java.nio.file.Path;

public final class FormulasExample implements Example {
    @Override
    public String name() {
        return "dev-ref/formulas";
    }

    @Override
    public String description() {
        return "Calculate formulas, dynamic arrays, localized formulas, and a Java UDF";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Formulas");
            workbook.setCalcMode(CalcMode.MANUAL);
            sheet.range("A2:A6").setNumbers(new double[] {3, 1, 3, 2, 1});
            sheet.cell(0, 0).setFormula("SORT(UNIQUE(A2:A6))");
            workbook.registerFunction("TRIANGLE", args -> args[0] * (args[0] + 1) / 2);
            sheet.cell(0, 2).setFormula("TRIANGLE(10)");
            sheet.cell(0, 3).setFormulaLocalized("SUMME(A2:A6)", "de-DE");
            workbook.recalculate();

            CellValue udf = sheet.value(0, 2);
            CellValue sum = workbook.evaluate("SUM(A2:A6)", sheet);
            System.out.printf("Functions: %d; XLOOKUP: %s; UDF: %s; sum: %s%n",
                    Functions.count(), Functions.contains("XLOOKUP"), udf, sum);
            workbook.save(outputDirectory.resolve("formulas.xlsx"));
        }
    }
}

