import io.keikai.axyra.sheets.Cell;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.ConditionalFormat;
import io.keikai.axyra.sheets.style.Color;
import java.nio.file.Path;

/** Preserves existing formats and adds a margin-based Excel rule. */
public final class FormatExistingReport {
    public static void main(String[] args) {
        Path output = Path.of("formatted-sales-highlighted.xlsx");
        try (Workbook wb = Workbook.open(Path.of("formatted-sales.xlsx"))) {
            Sheet sheet = wb.sheet("Sales");
            wb.registerNamedStyle("ReportCurrency", sheet.cell(1, 2).style());
            sheet.range("C2:C4").applyNamedStyle("ReportCurrency");
            for (int col = 0; col < 4; col++) {
                Cell cell = sheet.cell(2, col);
                cell.setStyle(cell.style().toBuilder()
                        .fill(Color.parse("#F2F6FC")).build());
            }
            sheet.range("D2:D4").addConditionalFormat(ConditionalFormat.builder()
                    .cellIs("lessThan", "0.18")
                    .fontColor(Color.parse("#9C0006")).bold()
                    .priority(1).build());
            wb.save(output);
        }
        try (Workbook check = Workbook.open(output)) {
            Sheet sheet = check.sheet("Sales");
            if (!"yyyy-mm-dd".equals(sheet.cell(2, 0).style().numberFormat())
                    || !"0.0%".equals(sheet.cell(2, 3).style().numberFormat())
                    || sheet.conditionalFormats().size() != 1
                    || !check.hasNamedStyle("ReportCurrency")) {
                throw new IllegalStateException("Expected preserved styles and margin rule");
            }
        }
        System.out.println("Wrote " + output);
    }
}
