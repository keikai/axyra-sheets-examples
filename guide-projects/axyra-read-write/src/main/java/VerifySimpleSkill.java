import io.keikai.axyra.sheets.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Path;
import java.util.*;

public final class VerifySimpleSkill {
    private static void check(boolean ok, String label) {
        if (!ok) throw new AssertionError(label);
    }
    public static void main(String[] args) throws Exception {
        var expected = List.of(new SimpleSkillWriter.Sale("00042", "Laptop", 1250),
                new SimpleSkillWriter.Sale("00043", "Mouse", 80));
        double expectedTotal = expected.stream().mapToDouble(SimpleSkillWriter.Sale::amount).sum();
        var report = SimpleSkillReader.read(Path.of("input/sales.xlsx"));
        check(report.orders().size() == expected.size(), "reader order count");
        for (int r = 0; r < expected.size(); r++) {
            var actual = report.orders().get(r); var row = expected.get(r);
            check(actual.orderId().equals(row.orderId()) && actual.product().equals(row.product())
                    && actual.amount() == row.amount(), "reader record " + r);
        }
        check(report.total() != null && report.total() == expectedTotal, "reader workbook total");
        Map<String, Object> evidence = new LinkedHashMap<>();
        evidence.put("reader", Map.of("expectedOrders", expected, "actualOrders", report.orders(),
                "expectedTotal", expectedTotal, "actualTotal", report.total()));
        try (Workbook workbook = Workbook.open(Path.of("output/ai-sales-simple.xlsx"))) {
            var sheet = workbook.sheet("Sales");
            for (int r = 0; r < expected.size(); r++) {
                var row = expected.get(r);
                check(sheet.cell(r + 1, 0).value().equals(CellValue.text(row.orderId())), "text ID");
                check(sheet.cell(r + 1, 1).value().equals(CellValue.text(row.product())), "product");
                check(sheet.cell(r + 1, 2).value().equals(CellValue.number(row.amount())), "numeric amount");
            }
            for (int c = 0; c < 3; c++) check(sheet.cell(0, c).style().bold(), "bold header");
            int totalRow = expected.size() + 1;
            String formula = "SUM(C2:C" + totalRow + ")";
            check(formula.equals(sheet.cell(totalRow, 2).formula()), "stored formula");
            double actualTotal = ((CellValue.Number)sheet.cell(totalRow, 2).value()).value();
            check(actualTotal == expectedTotal, "saved total");
            String display = sheet.cell(totalRow, 2).formattedText();
            check(display.equals(String.format(Locale.ROOT, "%,.2f", expectedTotal)), "total format");
            check(sheet.cell(1, 2).formattedText().equals("1,250.00"), "amount format");
            evidence.put("writer", Map.of("expectedTotal", expectedTotal, "actualTotal", actualTotal,
                    "formula", sheet.cell(totalRow, 2).formula(), "displayedTotal", display,
                    "textIds", true, "numericAmounts", true, "boldHeaders", true));
            System.out.printf(Locale.ROOT, "Reader total: expected %.2f; actual %.2f%n",
                    expectedTotal, report.total());
            System.out.printf(Locale.ROOT, "Saved total: expected %.2f; actual %.2f%n",
                    expectedTotal, actualTotal);
            System.out.println("Stored formula: " + formula + "; displayed total: " + display);
            System.out.println("PASS: text IDs, Java records, numeric amounts, bold headers, and formats");
        }
        new ObjectMapper().writerWithDefaultPrettyPrinter().writeValue(
                Path.of("ai-session-simple/expected-actual.json").toFile(), evidence);
    }
}
