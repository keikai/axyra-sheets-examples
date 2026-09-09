package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.StreamStyle;
import io.keikai.axyra.sheets.io.StreamWorkbook;

import java.nio.file.Path;
import java.time.LocalDate;

public final class StreamingExample implements Example {
    @Override
    public String name() {
        return "dev-ref/streaming";
    }

    @Override
    public String description() {
        return "Write rows forward-only and use the columnar workbook backend";
    }

    @Override
    public void run(Path outputDirectory) {
        Path streamed = outputDirectory.resolve("streamed.xlsx");
        try (StreamWorkbook output = StreamWorkbook.create(streamed)) {
            int header = output.addStyle(StreamStyle.defaults().withBold());
            int money = output.addStyle(StreamStyle.defaults().withNumberFormat("#,##0.00"));
            int date = output.addStyle(StreamStyle.defaults().withNumberFormat("yyyy-mm-dd"));
            output.startSheet("Data");
            output.row(0).addText("Date", header).addText("Region", header)
                    .addText("Amount", header).commit();
            for (int i = 0; i < 10; i++) {
                output.row(i + 1).addDate(LocalDate.of(2026, 9, 1).plusDays(i), date)
                        .addText(i % 2 == 0 ? "APAC" : "EMEA")
                        .addNumber(1000 + i * 25, money).commit();
            }
            output.finish();
        }

        try (Workbook workbook = Workbook.createColumnar()) {
            Sheet sheet = workbook.sheet(0);
            sheet.range("A1:A1000").setNumbers(new double[1000]);
            workbook.save(outputDirectory.resolve("columnar.xlsx"));
        }
    }
}

