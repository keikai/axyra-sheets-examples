package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.template.SmartMarkers;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

public final class GenerateReportExample implements Example {
    public record Line(String region, String product, int qty, double amount) {
    }

    @Override
    public String name() {
        return "guides/generate-report";
    }

    @Override
    public String description() {
        return "Fill an Excel template with Smart Markers and render a report";
    }

    @Override
    public void run(Path outputDirectory) {
        Path template = createTemplate(outputDirectory.resolve("sales-template.xlsx"));
        Map<String, Object> data = Map.of(
                "title", "Q3 2026",
                "generated", LocalDate.of(2026, 9, 9),
                "lines", List.of(
                        new Line("APAC", "Widget", 12, 1450),
                        new Line("EMEA", "Gadget", 3, 1100)));

        try (Workbook workbook = Workbook.open(template)) {
            Sheet sheet = workbook.sheet(0);
            SmartMarkers.process(sheet, data);
            workbook.recalculate();
            workbook.save(outputDirectory.resolve("sales-Q3-2026.xlsx"));
            workbook.renderPdf(outputDirectory.resolve("sales-Q3-2026.pdf"),
                    PdfOptions.builder().title("Sales Report Q3 2026")
                            .fitTo(PdfOptions.Fit.WIDTH).bookmarks(true).build());
        }
    }

    private static Path createTemplate(Path path) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Sales Report");
            sheet.range("A1:D4").setValues(new CellValue[][] {
                    {CellValue.text("${title}"), CellValue.text("${generated}"),
                            CellValue.blank(), CellValue.blank()},
                    {CellValue.text("Region"), CellValue.text("Product"),
                            CellValue.text("Qty"), CellValue.text("Amount")},
                    {CellValue.text("${lines.region}"), CellValue.text("${lines.product}"),
                            CellValue.text("${lines.qty}"), CellValue.text("${lines.amount}")},
                    {CellValue.text("Total"), CellValue.blank(), CellValue.blank(),
                            CellValue.text("${lines.amount:sum}")}
            });
            workbook.save(path);
        }
        return path;
    }
}
