package io.keikai.axyra.examples.guides;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetType;
import io.keikai.axyra.sheets.Visibility;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.PdfOptions;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/** The complete program from "How to Convert Excel to PDF in Java". */
public final class ExcelToPdfExample implements Example {
    @Override
    public String name() {
        return "guides/convert-excel-to-pdf";
    }

    @Override
    public String description() {
        return "Convert every visible worksheet to a fitted, bookmarked PDF";
    }

    @Override
    public void run(Path outputDirectory) {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("pdf-source.xlsx"));
        Path target = outputDirectory.resolve("report.pdf");

        try (Workbook wb = Workbook.open(source)) {
            List<String> worksheets = new ArrayList<>();

            for (int i = 0; i < wb.sheetCount(); i++) {
                Sheet sheet = wb.sheet(i);
                // Naming a hidden sheet would render it, so skip those too.
                if (sheet.sheetType() != SheetType.WORKSHEET
                        || sheet.visibility() != Visibility.VISIBLE) {
                    continue;
                }
                worksheets.add(sheet.name());

                sheet.setLandscape(true);
                sheet.setPaperSize(9);                 // A4
                sheet.setPageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
                sheet.setPrintTitleRows("$1:$1");      // header on every page
            }

            PdfOptions options = PdfOptions.builder()
                    .title("Quarterly Report")
                    .author("Finance")
                    .sheetNames(worksheets.toArray(new String[0]))
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .bookmarks(true)
                    .build();

            wb.renderPdf(target, options);

            String first = worksheets.get(0);
            int pages = wb.sheetPageCount(wb.sheet(first).index());
            System.out.println("wrote " + target + "; " + first + " is " + pages + " page(s)");
        }
    }
}
