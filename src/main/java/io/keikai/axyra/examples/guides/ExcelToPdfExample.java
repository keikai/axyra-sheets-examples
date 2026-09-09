package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetType;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

public final class ExcelToPdfExample implements Example {
    @Override
    public String name() {
        return "guides/convert-excel-to-pdf";
    }

    @Override
    public String description() {
        return "Convert an Excel workbook to a fitted, bookmarked PDF";
    }

    @Override
    public void run(Path outputDirectory) {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("pdf-source.xlsx"));
        Path target = outputDirectory.resolve("report.pdf");
        try (Workbook workbook = Workbook.open(source)) {
            List<String> worksheets = new ArrayList<>();
            for (int i = 0; i < workbook.sheetCount(); i++) {
                Sheet sheet = workbook.sheet(i);
                if (sheet.sheetType() != SheetType.WORKSHEET) continue;
                worksheets.add(sheet.name());
                sheet.setLandscape(true);
                sheet.setPaperSize(9);
                sheet.setPageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
                sheet.setPrintTitleRows("$1:$1");
            }
            PdfOptions options = PdfOptions.builder()
                    .title("Quarterly Report")
                    .author("Finance")
                    .sheetNames(worksheets.toArray(new String[0]))
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .bookmarks(true)
                    .build();
            workbook.renderPdf(target, options);
            System.out.println("Wrote " + target);
        }
    }
}

