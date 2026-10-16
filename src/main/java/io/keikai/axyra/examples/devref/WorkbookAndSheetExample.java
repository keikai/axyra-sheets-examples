package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.NamedRange;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Visibility;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

public final class WorkbookAndSheetExample implements Example {
    @Override
    public String name() {
        return "dev-ref/workbook-and-sheet";
    }

    @Override
    public String description() {
        return "Manage sheets, workbook settings, page layout, and named ranges";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet summary = workbook.createSheet("Summary");
            Sheet config = workbook.createSheet("Config");
            config.setNumber(1, 1, 0.05);

            summary.setTabColor(Color.rgb(0x2F, 0x55, 0x97));
            summary.setHeader("&C&\"Calibri,Bold\"Quarterly Report");
            summary.setFooter("&LPrinted &D&RPage &P of &N");
            summary.setRowHeight(0, 28);
            summary.setColumnWidth(0, 24);
            summary.freezePanes(1, 0);
            config.setVisibility(Visibility.VERY_HIDDEN);

            workbook.setCalcMode(CalcMode.MANUAL);
            workbook.setIterativeCalc(true, 100, 0.001);
            workbook.setDate1904(false);
            workbook.addName("TaxRate", "Config!$B$2");
            for (NamedRange name : workbook.names()) {
                System.out.println(name.name() + " = " + name.formula());
            }
            workbook.save(outputDirectory.resolve("workbook-and-sheet.xlsx"));
        }
    }
}

