package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;

import java.nio.file.Path;

public final class CreateExcelFileExample implements Example {
    @Override
    public String name() {
        return "guides/create-excel-file";
    }

    @Override
    public String description() {
        return "Create a styled sales workbook with formulas and frozen headers";
    }

    @Override
    public void run(Path outputDirectory) {
        Path output = outputDirectory.resolve("q3-sales.xlsx");
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Q3 Sales");
            sheet.range("A1:D1").setValues(new CellValue[][] {{
                    CellValue.text("Region"), CellValue.text("Q1"),
                    CellValue.text("Q2"), CellValue.text("Total")
            }});
            sheet.range("A2:C4").setValues(new CellValue[][] {
                    {CellValue.text("APAC"), CellValue.number(1200), CellValue.number(1450)},
                    {CellValue.text("EMEA"), CellValue.number(980), CellValue.number(1100)},
                    {CellValue.text("Americas"), CellValue.number(1610), CellValue.number(1720)}
            });
            for (int row = 2; row <= 4; row++) {
                sheet.cell(row - 1, 3).setFormula("B" + row + "+C" + row);
            }
            sheet.cell(4, 0).setValue(CellValue.text("Total"));
            sheet.cell(4, 1).setFormula("SUM(B2:B4)");
            sheet.range("B5:B5").autoFill(sheet.range("B5:D5"));
            workbook.recalculate();

            CellStyle header = CellStyle.builder().bold()
                    .fontColor(Color.parse("#FFFFFF"))
                    .fill(Color.parse("#2F5597"))
                    .horizontalAlign(HAlign.CENTER)
                    .allBorders(BorderStyle.THIN, Color.parse("#1F3864"))
                    .build();
            CellStyle money = CellStyle.builder().numberFormat("#,##0.00").build();
            CellStyle total = money.toBuilder().bold().build();
            sheet.range("A1:D1").setStyle(header);
            sheet.range("B2:D4").setStyle(money);
            sheet.range("A5:D5").setStyle(total);
            sheet.setColumnWidth(0, 14);
            for (int col = 1; col <= 3; col++) {
                sheet.autoFitColumn(col);
            }
            sheet.freezePanes(1, 0);
            workbook.save(output);
        }
        System.out.println("Wrote " + output);
    }
}
