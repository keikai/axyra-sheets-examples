import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.nio.file.*;
import java.util.*;

/** Independent reader: compares the fixture's complete authored region. */
public final class CompareReports {
    public static void main(String[] args) throws Exception {
        List<String> differences = new ArrayList<>();
        try (var a = new XSSFWorkbook(Files.newInputStream(Path.of("output/poi-sales.xlsx")));
             var b = new XSSFWorkbook(Files.newInputStream(Path.of("output/axyra-sales.xlsx")))) {
            XSSFSheet old = a.getSheet("Sales"), migrated = b.getSheet("Sales");
            for (XSSFWorkbook wb : List.of(a, b)) {
                for (int i = 0; i < wb.getNumberOfFonts(); i++) {
                    equal(differences, "explicit font scheme " + i, FontScheme.NONE,
                            wb.getFontAt(i).getScheme());
                }
            }
            equal(differences, "default font", font(a.getFontAt(0)), font(b.getFontAt(0)));
            compareScheme(differences, "default font", a.getFontAt(0), b.getFontAt(0));
            equal(differences, "row height", old.getDefaultRowHeightInPoints(),
                    migrated.getDefaultRowHeightInPoints());
            equal(differences, "merged cells", old.getMergedRegions(), migrated.getMergedRegions());
            equal(differences, "tables", old.getTables().size(), migrated.getTables().size());
            equal(differences, "freeze rows", old.getPaneInformation().getHorizontalSplitPosition(),
                    migrated.getPaneInformation().getHorizontalSplitPosition());
            equal(differences, "freeze cols", old.getPaneInformation().getVerticalSplitPosition(),
                    migrated.getPaneInformation().getVerticalSplitPosition());
            for (int c = 0; c < 4; c++) {
                equal(differences, "column width " + c, old.getColumnWidth(c), migrated.getColumnWidth(c));
            }
            for (int r = 0; r < 4; r++) {
                for (int c = 0; c < 4; c++) {
                    Cell x = old.getRow(r).getCell(c), y = migrated.getRow(r).getCell(c);
                    String label = "R" + (r + 1) + "C" + (c + 1);
                    equal(differences, label + " value/formula/cache", value(x), value(y));
                    equal(differences, label + " style", style(x), style(y));
                    if (x != null && y != null && x.getCellType() != CellType.BLANK
                            && y.getCellType() != CellType.BLANK) {
                        compareScheme(differences, label,
                                ((XSSFCellStyle)x.getCellStyle()).getFont(),
                                ((XSSFCellStyle)y.getCellStyle()).getFont());
                    }
                }
            }
            equal(differences, "original sheet names", List.of("Sales"), sheetNames(a));
            boolean evaluation = sheetNames(b).contains("Axyra Evaluation Copy");
            equal(differences, "migrated sheet names",
                    evaluation ? List.of("Axyra Evaluation Copy", "Sales") : List.of("Sales"), sheetNames(b));
            System.out.println("POI tabs: " + sheetNames(a));
            System.out.println("Axyra tabs: " + sheetNames(b));
            System.out.println("Sales A1:D4, defaults, widths, freeze panes, merges and tables checked.");
            if (!differences.isEmpty()) {
                differences.forEach(System.out::println);
                throw new AssertionError(differences.size() + " differences");
            }
            System.out.println("PASS: values and cell types");
            System.out.println("PASS: formula text and calculated results");
            System.out.println("PASS: number formats and checked styles, including font schemes");
            System.out.println("PASS: checked worksheet structure");
            if (evaluation) System.out.println("Expected difference: Axyra evaluation notice tab/rows are retained.");
            System.out.println("Not compared: ZIP bytes, package metadata, desktop Excel layout, PDF output.");
        }
    }

    private static void compareScheme(List<String> differences,
                                      String label, XSSFFont old, XSSFFont migrated) {
        equal(differences, label + " font scheme", old.getScheme(), migrated.getScheme());
    }

    private static List<String> sheetNames(XSSFWorkbook wb) {
        List<String> names = new ArrayList<>();
        for (Sheet s : wb) names.add(s.getSheetName());
        return names;
    }

    private static String value(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return "blank";
        return switch (cell.getCellType()) {
            case STRING -> "text:" + cell.getStringCellValue();
            case NUMERIC -> "number:" + cell.getNumericCellValue();
            case FORMULA -> "formula:" + cell.getCellFormula() + ":"
                    + cell.getCachedFormulaResultType() + ":" + cell.getNumericCellValue();
            default -> throw new AssertionError("Unexpected type " + cell.getCellType());
        };
    }

    private static String style(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return "blank";
        XSSFCellStyle s = (XSSFCellStyle) cell.getCellStyle();
        return List.of(s.getDataFormatString(), font(s.getFont()),
                s.getFillPattern(), s.getAlignment(), s.getVerticalAlignment(),
                s.getWrapText(), s.getBorderTop(), s.getBorderBottom(),
                s.getBorderLeft(), s.getBorderRight()).toString();
    }

    private static String font(XSSFFont f) {
        XSSFColor color = f.getXSSFColor();
        String col = color == null || color.isAuto() ? "auto"
                : color.isThemed() ? "theme:" + color.getTheme()
                : color.isRGB() ? "rgb:" + color.getARGBHex()
                : "indexed:" + color.getIndexed();
        return List.of(f.getFontName(), f.getFontHeightInPoints(), f.getBold(),
                f.getItalic(), col, f.getFamily()).toString();
    }

    private static void equal(List<String> differences, String label, Object a, Object b) {
        if (!Objects.equals(a, b)) differences.add(label + ": " + a + " -> " + b);
    }
}
