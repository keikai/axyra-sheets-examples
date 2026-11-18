import io.keikai.axyra.sheets.*;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.List;

/** Focused regression checks for TSV fields and saved database report formats. */
public final class VerifyRevisions {
    private static void check(boolean ok, String label) {
        if (!ok) throw new AssertionError(label);
    }
    public static void main(String[] args) throws Exception {
        String leading = ReadAllWorksheets.toTsv(new CellValue[]{
                CellValue.blank(), CellValue.text("Laptop"), CellValue.number(1250)});
        check(leading.equals("\tLaptop\t1250.0"), "leading blank TSV field");
        check(ReadAllWorksheets.toTsv(new CellValue[]{CellValue.blank(),
                CellValue.blank(), CellValue.blank()}).equals("\t\t"), "all-blank TSV row");
        check(ReadAllWorksheets.toTsv(new CellValue[]{CellValue.text("a\tb\nc"),
                CellValue.blank()}).equals("a b c\t"), "TSV control characters/trailing blank");
        Files.createDirectories(Path.of("output"));
        var rows = List.of(new DatabaseSalesReport.Sale("East", "Laptop",
                        LocalDate.of(2026, 10, 27), 1250.5),
                new DatabaseSalesReport.Sale("West", "Mouse", null, null));
        DatabaseSalesReport.writeReport(rows, Path.of("output/q3-sales.xlsx"));
        try (Workbook wb = Workbook.open(Path.of("output/q3-sales.xlsx"))) {
            Sheet s = wb.sheet("Q3 Sales");
            check(s.cell(1, 2).formattedText().equals("2026-10-27"), "saved date format");
            check(s.cell(1, 2).dateValue().equals(LocalDate.of(2026, 10, 27)), "saved date value");
            check(s.cell(1, 3).formattedText().equals("1,250.50"), "saved amount format");
            check(s.cell(1, 3).value() instanceof CellValue.Number, "numeric amount");
            check(s.cell(2, 2).value() instanceof CellValue.Blank, "null date");
            check(s.cell(2, 3).value() instanceof CellValue.Blank, "null amount");
        }
        DatabaseSalesReport.writeReport(List.of(), Path.of("output/q3-sales-empty.xlsx"));
        try (Workbook wb = Workbook.open(Path.of("output/q3-sales-empty.xlsx"))) {
            check(wb.sheet("Q3 Sales").cell(0, 0).formattedText().equals("Region"), "empty export header");
        }
        WriteJsonSalesReport.main(new String[0]);
        try (Workbook wb = Workbook.open(Path.of("q3-json-sales.xlsx"))) {
            Sheet s = wb.sheet("Q3 Sales");
            check(s.cell(1, 2).dateValue().equals(LocalDate.of(2026, 10, 27)), "JSON LocalDate");
            check(s.cell(1, 2).formattedText().equals("2026-10-27"), "JSON date format");
            check(s.cell(1, 3).formattedText().equals("1,250.50"), "JSON amount format");
        }
        System.out.println("PASS: leading/all/trailing blank TSV fields; escaped TSV controls;");
        System.out.println("saved date and numeric formats; null date/amount; empty export headers.");
        System.out.println("PASS: JSON LocalDate deserialization and saved date/amount formats.");
        System.out.println("Database writeReport tested with Java records; no live JDBC server used.");
    }
}
