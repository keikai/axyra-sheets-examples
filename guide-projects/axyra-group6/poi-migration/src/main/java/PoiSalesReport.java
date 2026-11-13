import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.nio.file.*;
import java.time.LocalDate;

public final class PoiSalesReport {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("output"));
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Sales");
            sheet.setDefaultRowHeightInPoints(15);
            String[] labels = {"Order ID", "Date", "Amount", "With Tax"};
            Row header = sheet.createRow(0);
            Font bold = wb.createFont();
            bold.setFontName("Calibri");
            bold.setFontHeightInPoints((short) 11);
            bold.setBold(true);
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFont(bold);
            CellStyle date = wb.createCellStyle();
            date.setDataFormat(wb.createDataFormat().getFormat("yyyy-mm-dd"));
            CellStyle money = wb.createCellStyle();
            money.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
            for (int c = 0; c < labels.length; c++) {
                header.createCell(c).setCellValue(labels[c]);
                header.getCell(c).setCellStyle(headerStyle);
                sheet.setColumnWidth(c, 20 * 256);
            }
            String[] ids = {"00042", "00043"};
            double[] amounts = {1250, 80};
            for (int i = 0; i < ids.length; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(ids[i]);
                row.createCell(1).setCellValue(
                        LocalDate.of(2026, 10, 27 + i).atStartOfDay());
                row.getCell(1).setCellStyle(date);
                row.createCell(2).setCellValue(amounts[i]);
                row.getCell(2).setCellStyle(money);
                row.createCell(3).setCellFormula("C" + (i + 2) + "*1.2");
                row.getCell(3).setCellStyle(money);
            }
            Row total = sheet.createRow(3);
            total.createCell(0).setCellValue("Total");
            total.createCell(2).setCellFormula("SUM(C2:C3)");
            total.createCell(3).setCellFormula("SUM(D2:D3)");
            total.getCell(2).setCellStyle(money);
            total.getCell(3).setCellStyle(money);
            sheet.createFreezePane(0, 1);
            wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
            try (var out = Files.newOutputStream(Path.of("output/poi-sales.xlsx"))) {
                wb.write(out);
            }
        }
    }
}
