package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.template.TemplateMarkers;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

public final class TemplateMarkersExample implements Example {
    record Line(String item, int qty, double amount) {
    }

    @Override
    public String name() {
        return "dev-ref/template-markers";
    }

    @Override
    public String description() {
        return "Fill scalar, list, and aggregate Template Markers";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Invoice");
            sheet.range("A1:C4").setValues(new CellValue[][] {
                    {CellValue.text("Invoice for ${customer.name}"), CellValue.blank(), CellValue.blank()},
                    {CellValue.text("Item"), CellValue.text("Qty"), CellValue.text("Amount")},
                    {CellValue.text("${lines.item}"), CellValue.text("${lines.qty}"), CellValue.text("${lines.amount}")},
                    {CellValue.text("Total"), CellValue.blank(), CellValue.text("${lines.amount:sum}")}
            });
            Map<String, Object> data = Map.of(
                    "customer", Map.of("name", "Acme Corp"),
                    "lines", List.of(new Line("Widget", 12, 51), new Line("Gadget", 3, 27)));
            TemplateMarkers.process(sheet, data);
            workbook.recalculateDirty();
            workbook.save(outputDirectory.resolve("template-markers.xlsx"));
        }
    }
}

