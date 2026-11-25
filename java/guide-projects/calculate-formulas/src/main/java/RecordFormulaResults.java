import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

/** Records the sales totals before and after the quantity correction. */
public final class RecordFormulaResults {
    public static void main(String[] args) throws Exception {
        Path out = Path.of("output/results");
        Files.createDirectories(out);
        try (Workbook wb = Workbook.open(Path.of("output/sales-calculated.xlsx"))) {
            wb.setCalcMode(CalcMode.AUTOMATIC);
            Sheet sales = wb.sheet("Sales");
            sales.cell(1, 2).setValue(CellValue.number(2));
            double before = ((CellValue.Number) wb.sheet("Summary").cell(1, 8).value()).value();
            wb.save(out.resolve("formula-before.xlsx"));
            sales.cell(1, 2).setValue(CellValue.number(4));
            double after = ((CellValue.Number) wb.sheet("Summary").cell(1, 8).value()).value();
            if (before != 445 || after != 495) throw new AssertionError("Wrong formula preview totals");
            new ObjectMapper().writeValue(out.resolve("formula-evidence.json").toFile(),
                    Map.of("beforeQuantity", 2, "afterQuantity", 4, "unitPrice", 25,
                            "beforeTotal", before, "afterTotal", after));
        }
        System.out.println("Wrote formula results to " + out);
    }
}
