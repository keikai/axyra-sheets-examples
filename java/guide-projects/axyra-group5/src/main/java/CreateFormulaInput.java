import io.keikai.axyra.sheets.*;
import io.keikai.axyra.sheets.style.*;
import java.nio.file.*;

public final class CreateFormulaInput {
    public static void main(String[] args) throws Exception {
        Files.createDirectories(Path.of("input"));
        try (Workbook wb = Workbook.create()) {
            Sheet products = wb.createSheet("Products");
            products.range("A1:B4").setValues(new CellValue[][] {
                {CellValue.text("SKU"), CellValue.text("Price")},
                {CellValue.text("P100"), CellValue.number(25)},
                {CellValue.text("P200"), CellValue.number(40)},
                {CellValue.text("P300"), CellValue.number(15)}
            });
            Sheet sales = wb.createSheet("Sales");
            sales.range("A1:E1").setValues(new CellValue[][] { {
                CellValue.text("Region"), CellValue.text("SKU"),
                CellValue.text("Qty"), CellValue.text("Unit price"),
                CellValue.text("Amount")
            }});
            sales.range("A2:C5").setValues(new CellValue[][] {
                {CellValue.text("North"), CellValue.text("P100"), CellValue.number(2)},
                {CellValue.text("South"), CellValue.text("P200"), CellValue.number(3)},
                {CellValue.text("North"), CellValue.text("P300"), CellValue.number(5)},
                {CellValue.text("East"), CellValue.text("P100"), CellValue.number(8)}
            });
            CellStyle header = CellStyle.builder().bold()
                    .fontColor(Color.parse("#FFFFFF"))
                    .fill(Color.parse("#1F3B68")).build();
            products.range("A1:B1").setStyle(header);
            sales.range("A1:E1").setStyle(header);
            sales.range("D2:E5").setStyle(CellStyle.builder()
                    .numberFormat("#,##0.00").build());
            for (int c = 0; c < 5; c++) sales.setColumnWidth(c, 16);
            wb.save(Path.of("input/sales-input.xlsx"));
        }
    }
}
