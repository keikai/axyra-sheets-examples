import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;
import io.keikai.axyra.sheets.style.Side;
import io.keikai.axyra.sheets.style.VAlign;

import java.nio.file.Path;
import java.time.LocalDate;

public final class FormatExcelReport {
    public static void main(String[] args) {
        Path output = Path.of("formatted-sales.xlsx");

        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Sales");
            sheet.range("A1:D4").setValues(new CellValue[][] {
                {CellValue.text("Order date"), CellValue.text("Region"),
                    CellValue.text("Revenue (USD)"), CellValue.text("Margin")},
                {CellValue.blank(), CellValue.text("North"),
                    CellValue.number(12500.5), CellValue.number(0.185)},
                {CellValue.blank(), CellValue.text("South"),
                    CellValue.number(9800), CellValue.number(0.152)},
                {CellValue.blank(), CellValue.text("East"),
                    CellValue.number(14200.75), CellValue.number(0.21)}
            });
            sheet.cell(1, 0).setDate(LocalDate.of(2026, 10, 1));
            sheet.cell(2, 0).setDate(LocalDate.of(2026, 10, 2));
            sheet.cell(3, 0).setDate(LocalDate.of(2026, 10, 3));

            Color grid = Color.parse("#D9E2F3");
            CellStyle body = CellStyle.builder()
                    .fontName("Calibri").fontSize(11)
                    .verticalAlign(VAlign.CENTER)
                    .allBorders(BorderStyle.THIN, grid)
                    .build();
            CellStyle header = body.toBuilder()
                    .bold()
                    .fontColor(Color.parse("#FFFFFF"))
                    .fill(Color.parse("#1F3B68"))
                    .horizontalAlign(HAlign.CENTER)
                    .wrapText(true)
                    .build();
            CellStyle date = body.toBuilder()
                    .numberFormat("yyyy-mm-dd").build();
            CellStyle money = body.toBuilder()
                    .numberFormat("\"$\"#,##0.00;[Red](\"$\"#,##0.00)")
                    .horizontalAlign(HAlign.RIGHT).build();
            CellStyle percent = body.toBuilder()
                    .numberFormat("0.0%")
                    .horizontalAlign(HAlign.RIGHT).build();

            sheet.range("A1:D1").setStyle(header);
            sheet.range("A2:D4").setStyle(body);
            sheet.range("A2:A4").setStyle(date);
            sheet.range("C2:C4").setStyle(money);
            sheet.range("D2:D4").setStyle(percent);

            sheet.cell(4, 1).setValue(CellValue.text("Total"));
            sheet.cell(4, 2).setFormula("SUM(C2:C4)");
            CellStyle total = body.toBuilder()
                    .bold().fill(Color.parse("#EAF0F8"))
                    .border(Side.TOP, BorderStyle.DOUBLE,
                            Color.parse("#1F3B68"))
                    .build();
            sheet.range("A5:D5").setStyle(total);
            sheet.cell(4, 2).setStyle(total.toBuilder()
                    .numberFormat(money.numberFormat())
                    .horizontalAlign(HAlign.RIGHT).build());

            sheet.setRowHeight(0, 32);
            for (int row = 1; row <= 4; row++) {
                sheet.setRowHeight(row, 22);
            }
            sheet.setColumnWidth(0, 16);
            sheet.setColumnWidth(1, 16);
            sheet.setColumnWidth(2, 22);
            sheet.setColumnWidth(3, 14);
            sheet.freezePanes(1, 0);
            wb.save(output);
        }

        try (Workbook check = Workbook.open(output)) {
            Sheet sales = check.sheet("Sales");
            if (!"0.0%".equals(sales.cell(1, 3).style().numberFormat())
                    || !sales.cell(4, 2).style().bold()) {
                throw new IllegalStateException("Expected formatting was lost");
            }
            System.out.println(sales.cell(1, 2).formattedText());
            System.out.println(sales.cell(1, 3).formattedText());
        }
        System.out.println("Wrote " + output.toAbsolutePath());
    }
}
