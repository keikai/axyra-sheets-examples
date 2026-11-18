import io.keikai.axyra.sheets.*;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.*;
import java.util.*;

/** Checks the new AI-generated examples and exports actual data for guide figures. */
public final class VerifySkillSession {
    private static void check(boolean ok, String message) {
        if (!ok) throw new AssertionError(message);
    }
    private static List<List<String>> preview(Sheet sheet) {
        List<List<String>> rows = new ArrayList<>();
        for (int r = 0; r < 4; r++) {
            List<String> row = new ArrayList<>();
            for (int c = 0; c < 3; c++) row.add(sheet.cell(r, c).formattedText());
            rows.add(row);
        }
        return rows;
    }
    public static void main(String[] args) throws Exception {
        Map<String, Object> evidence = new LinkedHashMap<>();
        try (Workbook wb = Workbook.open(Path.of("input/sales.xlsx"))) {
            evidence.put("input", preview(wb.sheet("Sales")));
        }
        try (Workbook wb = Workbook.open(Path.of("output/ai-sales.xlsx"))) {
            Sheet sheet = wb.sheet("Sales");
            check(sheet.cell(1, 0).value().equals(CellValue.text("00042")), "text ID 00042");
            check(sheet.cell(2, 0).value().equals(CellValue.text("00043")), "text ID 00043");
            check(sheet.cell(1, 2).value().equals(CellValue.number(1250)), "numeric 1250");
            check(sheet.cell(2, 2).value().equals(CellValue.number(80)), "numeric 80");
            for (int c = 0; c < 3; c++) check(sheet.cell(0, c).style().bold(), "bold header");
            check("SUM(C2:C3)".equals(sheet.cell(3, 2).formula()), "stored SUM formula");
            check(sheet.cell(3, 2).value().equals(CellValue.number(1330)), "saved numeric total");
            check(sheet.cell(3, 2).formattedText().equals("1,330.00"), "displayed total");
            check(sheet.cell(1, 2).formattedText().equals("1,250.00"), "displayed amount");
            evidence.put("output", preview(sheet));
            sheet.cell(1, 2).setValue(CellValue.number(1500));
            check(sheet.cell(3, 2).value().equals(CellValue.number(1580)), "dependent recalculation");
            wb.save(Path.of("output/ai-sales-updated.xlsx"));
        }
        try (Workbook wb = Workbook.open(Path.of("output/ai-sales-updated.xlsx"))) {
            check(wb.sheet("Sales").cell(3, 2).value().equals(CellValue.number(1580)), "reopened updated total");
        }
        SkillWriteSales.write(List.of(), Path.of("output/ai-sales-empty.xlsx"));
        try (Workbook wb = Workbook.open(Path.of("output/ai-sales-empty.xlsx"))) {
            check(wb.sheet("Sales").cell(1, 2).value().equals(CellValue.number(0)), "empty-list total");
        }
        new ObjectMapper().writerWithDefaultPrettyPrinter().writeValue(
                Path.of("ai-session/figure-evidence.json").toFile(), evidence);
        System.out.println("PASS: text IDs 00042/00043; numeric amounts 1250/80; bold headers");
        System.out.println("PASS: C4 formula SUM(C2:C3); numeric total 1330; displayed total 1,330.00");
        System.out.println("PASS: edit C2 to 1500; saved and reopened total 1580; empty-list total 0");
    }
}
