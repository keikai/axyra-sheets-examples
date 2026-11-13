import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import java.nio.file.*;
import java.time.LocalDate;

public final class AxyraSalesReport {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("output"));
        try (Workbook wb = Workbook.create()) {
            // Match the default font observed in the POI baseline's styles.xml.
            wb.setDefaultFont(wb.defaultStyle().toBuilder()
                    .fontName("Calibri").fontSize(11).fontColor(Color.indexed(8))
                    .fontScheme("minor").build());
            Sheet sheet = wb.createSheet("Sales");
            sheet.setDefaultRowHeight(15);
            sheet.range("A1:D1").setValues(new CellValue[][] { {
                CellValue.text("Order ID"), CellValue.text("Date"),
                CellValue.text("Amount"), CellValue.text("With Tax")
            }});
            CellStyle header = wb.defaultStyle().toBuilder().bold()
                    .fontColor(Color.auto()).fontFamily(0).fontScheme("none").build();
            CellStyle date = wb.defaultStyle().toBuilder()
                    .numberFormat("yyyy-mm-dd").build();
            CellStyle money = wb.defaultStyle().toBuilder()
                    .numberFormat("#,##0.00").build();
            sheet.range("A1:D1").setStyle(header);
            String[] ids = {"00042", "00043"};
            double[] amounts = {1250, 80};
            for (int i = 0; i < ids.length; i++) {
                int row = i + 1;
                sheet.cell(row, 0).setValue(CellValue.text(ids[i]));
                sheet.cell(row, 1).setDate(LocalDate.of(2026, 10, 27 + i));
                sheet.cell(row, 1).setStyle(date);
                sheet.cell(row, 2).setValue(CellValue.number(amounts[i]));
                sheet.cell(row, 3).setFormula("C" + (i + 2) + "*1.2");
                sheet.range(row, 2, row, 3).setStyle(money);
            }
            sheet.cell(3, 0).setValue(CellValue.text("Total"));
            sheet.cell(3, 2).setFormula("SUM(C2:C3)");
            sheet.cell(3, 3).setFormula("SUM(D2:D3)");
            sheet.range("C4:D4").setStyle(money);
            for (int c = 0; c < 4; c++) sheet.setColumnWidth(c, 20);
            sheet.freezePanes(1, 0);
            sheet.setActiveCell("A2", null);
            wb.recalculate();
            wb.save(Path.of("output/axyra-sales.xlsx"));
        }
    }
}
