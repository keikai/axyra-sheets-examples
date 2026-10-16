package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.format.Formats;
import io.keikai.axyra.sheets.format.ParsedValue;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;
import io.keikai.axyra.sheets.style.Side;
import io.keikai.axyra.sheets.style.Theme;
import io.keikai.axyra.sheets.style.VAlign;

import java.nio.file.Path;

public final class StylesExample implements Example {
    @Override
    public String name() {
        return "dev-ref/styles";
    }

    @Override
    public String description() {
        return "Build, serialize, register, and apply styles and number formats";
    }

    @Override
    public void run(Path outputDirectory) {
        CellStyle header = CellStyle.builder()
                .fontName("Calibri").fontSize(11).bold(true)
                .fontColor(Color.rgb(255, 255, 255)).fill(Color.parse("#2C3E50"))
                .horizontalAlign(HAlign.CENTER).verticalAlign(VAlign.CENTER)
                .wrapText(true)
                .border(Side.BOTTOM, BorderStyle.DOUBLE, Color.parse("#333333"))
                .build();
        CellStyle restored = CellStyle.fromJson(header.toJson());
        CellStyle money = CellStyle.builder().numberFormat("#,##0.00;[Red](#,##0.00)").build();

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Styles");
            sheet.cell(0, 0).setValue(CellValue.text("Amount"));
            sheet.setNumber(1, 0, 1234.5);
            sheet.cell(0, 0).setStyle(restored);
            workbook.registerNamedStyle("Currency", money);
            sheet.range("A2:A2").applyNamedStyle("Currency");

            Theme theme = Theme.builder().majorFont("Georgia").minorFont("Calibri")
                    .color(Theme.Slot.ACCENT1, Color.parse("#2C3E50"))
                    .build();
            workbook.setTheme(theme);
            ParsedValue parsed = Formats.parse("31/12/2026");
            System.out.println(Formats.format(1234.5, "#,##0.00") + "; parsed=" + parsed);
            workbook.save(outputDirectory.resolve("styles.xlsx"));
        }
    }
}
