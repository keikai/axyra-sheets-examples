package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.Cell;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.ConditionalFormat;
import io.keikai.axyra.sheets.content.Hyperlink;
import io.keikai.axyra.sheets.content.RichText;
import io.keikai.axyra.sheets.content.Validation;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

public final class ContentObjectsExample implements Example {
    @Override
    public String name() {
        return "dev-ref/content-objects";
    }

    @Override
    public String description() {
        return "Create a table, validation, conditional formatting, comments, and links";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.sheet(0);
            sheet.range("A1:D4").setValues(new CellValue[][] {
                    {CellValue.text("Region"), CellValue.text("Q1"),
                            CellValue.text("Q2"), CellValue.text("Total")},
                    {CellValue.text("APAC"), CellValue.number(1200),
                            CellValue.number(1450), CellValue.number(2650)},
                    {CellValue.text("EMEA"), CellValue.number(980),
                            CellValue.number(1100), CellValue.number(2080)},
                    {CellValue.text("Americas"), CellValue.number(1610),
                            CellValue.number(1720), CellValue.number(3330)}
            });
            sheet.range("A1:D4").createTable("Sales", "Region", "Q1", "Q2", "Total");

            Range totals = sheet.range("D2:D4");
            totals.addConditionalFormat(ConditionalFormat.builder()
                    .cellIs("greaterThan", "2500")
                    .fillColor(Color.parse("#FFF3CD"))
                    .fontColor(Color.parse("#8A6D3B"))
                    .bold().priority(1).build());
            sheet.range("A2:A20").setValidation(Validation.list("APAC", "EMEA", "Americas"));

            Cell note = sheet.cell(1, 3);
            note.setComment("Finance", "Adjusted after audit");
            note.setHyperlink(Hyperlink.url("https://keikai.io", "Keikai"));
            sheet.cell(5, 0).setRichText(RichText.builder()
                    .run("Total: ")
                    .run("reviewed", RichText.run().bold().color(Color.parse("#C0392B")))
                    .build());

            sheet.protect("pa55w0rd");
            sheet.addProtectedRange("Editors", "B2:D4", "range-pw");
            workbook.protectStructure("pa55w0rd");
            workbook.save(outputDirectory.resolve("content-objects.xlsx"));
        }
    }
}

