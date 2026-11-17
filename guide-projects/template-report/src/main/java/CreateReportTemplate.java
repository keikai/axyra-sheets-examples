import io.keikai.axyra.sheets.*;
import io.keikai.axyra.sheets.style.*;
import java.nio.file.*;

public final class CreateReportTemplate {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("input"));
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Invoice");
            sheet.cell(0, 0).setValue(CellValue.text("SALES REPORT"));
            sheet.cell(1, 0).setValue(CellValue.text("Customer"));
            sheet.cell(1, 1).setValue(CellValue.text("${customer}"));
            sheet.cell(2, 0).setValue(CellValue.text("Period"));
            sheet.cell(2, 1).setValue(CellValue.text("${period}"));
            sheet.range("A4:E4").setValues(new CellValue[][] { {
                CellValue.text("SKU"), CellValue.text("Product"),
                CellValue.text("Qty"), CellValue.text("Unit price (USD)"),
                CellValue.text("Amount (USD)")
            }});
            sheet.range("A5:D5").setValues(new CellValue[][] { {
                CellValue.text("${lines.sku}"), CellValue.text("${lines.product}"),
                CellValue.text("${lines.qty}"), CellValue.text("${lines.unitPrice}")
            }});
            sheet.cell(4, 4).setFormula("C5*D5");
            sheet.cell(5, 3).setValue(CellValue.text("Total (USD)"));
            sheet.cell(5, 4).setValue(CellValue.text("${lines.amount:sum}"));
            CellStyle body = CellStyle.builder().fontName("Calibri").fontSize(11)
                    .allBorders(BorderStyle.THIN, Color.parse("#D9E2F3")).build();
            sheet.range("A5:E5").setStyle(body);
            sheet.range("D5:E5").setStyle(body.toBuilder()
                    .numberFormat("#,##0.00").build());
            sheet.range("A4:E4").setStyle(body.toBuilder().bold()
                    .fill(Color.parse("#1F3B68"))
                    .fontColor(Color.parse("#FFFFFF")).wrapText(true).build());
            sheet.range("D6:E6").setStyle(body.toBuilder().bold()
                    .numberFormat("#,##0.00")
                    .fill(Color.parse("#EAF0F8")).build());
            sheet.cell(0, 0).setStyle(CellStyle.builder().bold().fontSize(18).build());
            double[] widths = {14, 24, 10, 23, 23};
            for (int c = 0; c < widths.length; c++) sheet.setColumnWidth(c, widths[c]);
            sheet.setRowHeight(3, 32);
            sheet.freezePanes(4, 0);
            wb.save(Path.of("input/sales-template.xlsx"));
        }
    }
}
