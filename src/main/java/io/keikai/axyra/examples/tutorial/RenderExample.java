package io.keikai.axyra.examples.tutorial;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.HtmlOptions;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.io.PdfOptions;

import java.nio.file.Files;
import java.nio.file.Path;

public final class RenderExample implements Example {
    @Override
    public String name() {
        return "tutorial/render";
    }

    @Override
    public String description() {
        return "Render a workbook to PDF, PNG, SVG, and HTML";
    }

    @Override
    public void run(Path outputDirectory) throws Exception {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("render-source.xlsx"));
        try (Workbook workbook = Workbook.open(source)) {
            PdfOptions pdf = PdfOptions.builder()
                    .title("Q3 Report")
                    .author("Finance")
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .bookmarks(true)
                    .build();
            workbook.renderPdf(outputDirectory.resolve("report.pdf"), pdf);

            ImageOptions image = ImageOptions.builder().scale(2).gridLines(true).build();
            int sales = workbook.sheet("Q3 Sales").index();   // look the sheet up by name
            Files.write(outputDirectory.resolve("report.png"), workbook.renderSheetPng(sales, image));
            Files.write(outputDirectory.resolve("report.svg"), workbook.renderSheetSvg(sales, image));
            workbook.saveHtml(outputDirectory.resolve("report.html"),
                    HtmlOptions.create().singleFile(true));
            System.out.println("Q3 Sales is " + workbook.sheetPageCount(sales) + " page(s)");
        }
    }
}

