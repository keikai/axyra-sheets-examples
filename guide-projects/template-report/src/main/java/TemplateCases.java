import java.nio.file.Path;
import java.util.List;

public final class TemplateCases {
    public static void main(String[] args) {
        Path input = Path.of("input/sales-template.xlsx");
        Path one = Path.of("output/one-line-sales.xlsx");
        var lines = List.of(FillReportTemplate.Line.of("00042", "Keyboard", 2, 49.5));
        FillReportTemplate.fill(input, one, lines);
        FillReportTemplate.verify(one, lines);
        try {
            FillReportTemplate.fill(input, Path.of("output/empty.xlsx"), List.of());
            throw new AssertionError("Empty input should have been rejected");
        } catch (IllegalArgumentException expected) {
            System.out.println("PASS: one-row report; empty input rejected");
        }
    }
}
