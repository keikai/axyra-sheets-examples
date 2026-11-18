package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.template.TemplateMarkers;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * The complete program from "How to Create Excel Reports Automatically in Java".
 * It builds the template itself so that it runs without any input files.
 */
public final class GenerateReportExample implements Example {
    public record Line(String region, String product, int qty, double amount) {
    }

    /** A 1x1 PNG standing in for assets/logo.png. */
    private static final byte[] LOGO = Base64.getDecoder().decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    @Override
    public String name() {
        return "guides/generate-report";
    }

    @Override
    public String description() {
        return "Fill an Excel template with Template Markers and render a report";
    }

    @Override
    public void run(Path outputDirectory) {
        Path template = createTemplate(outputDirectory.resolve("sales-template.xlsx"));
        String quarter = "Q3-2026";
        List<Line> lines = List.of(
                new Line("APAC", "Widget", 12, 51.00),
                new Line("APAC", "Gadget", 3, 27.00),
                new Line("EMEA", "Widget", 8, 34.00));

        Map<String, Object> data = new HashMap<>();
        data.put("title", quarter);
        data.put("generated", LocalDate.of(2026, 9, 9));
        data.put("lines", lines);
        data.put("logo", LOGO);

        Path xlsx = outputDirectory.resolve("sales-" + quarter + ".xlsx");
        Path pdf = outputDirectory.resolve("sales-" + quarter + ".pdf");

        try (Workbook wb = Workbook.open(template)) {
            Sheet sheet = wb.sheet("Sales Report");   // by name, not index

            TemplateMarkers.process(sheet, data);
            wb.recalculate();   // refresh template formulas that process did not touch

            sheet.setPrintTitleRows("$1:$2");   // title + header on every page
            wb.save(xlsx);

            wb.renderPdf(pdf, PdfOptions.builder()
                    .title("Sales Report " + quarter)
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .bookmarks(true)
                    .build());

            System.out.println("Total = " + sheet.cell(5, 3).formattedText());   // 112.00
        }
        System.out.println("Wrote " + xlsx + " and " + pdf);
    }

    /** The template from the guide: title row, header row, template row, total row. */
    private static Path createTemplate(Path path) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Sales Report");
            sheet.range("A1:D4").setValues(new CellValue[][] {
                    {CellValue.text("Sales Report — ${title}"), CellValue.blank(),
                            CellValue.text("${generated}"), CellValue.text("${image:logo}")},
                    {CellValue.text("Region"), CellValue.text("Product"),
                            CellValue.text("Qty"), CellValue.text("Amount")},
                    {CellValue.text("${lines.region}"), CellValue.text("${lines.product}"),
                            CellValue.text("${lines.qty}"), CellValue.text("${lines.amount}")},
                    {CellValue.text("Total"), CellValue.blank(), CellValue.blank(),
                            CellValue.text("${lines.amount:sum}")}
            });
            // Formatting on the template row is copied to every inserted row.
            CellStyle money = CellStyle.builder().numberFormat("#,##0.00").build();
            sheet.range("D3:D4").setStyle(money);
            sheet.range("A2:D2").setStyle(CellStyle.builder().bold().build());
            workbook.save(path);
        }
        return path;
    }
}
