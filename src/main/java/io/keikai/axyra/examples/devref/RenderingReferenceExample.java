package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.HtmlOptions;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.io.Renderer;

import java.nio.file.Files;
import java.nio.file.Path;

public final class RenderingReferenceExample implements Example {
    @Override
    public String name() {
        return "dev-ref/rendering";
    }

    @Override
    public String description() {
        return "Exercise PDF options and all documented image/HTML render paths";
    }

    @Override
    public void run(Path outputDirectory) throws Exception {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("rendering-source.xlsx"));
        try (Workbook workbook = Workbook.open(source)) {
            PdfOptions pdf = PdfOptions.builder()
                    .title("Axyra rendering example")
                    .watermark("DRAFT")
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .build();
            workbook.renderPdf(outputDirectory.resolve("rendering.pdf"), pdf);

            ImageOptions image = ImageOptions.builder().targetWidthPx(1200).gridLines(true).build();
            Files.write(outputDirectory.resolve("rendering.png"), workbook.renderSheetPng(0, image));
            Files.write(outputDirectory.resolve("rendering.svg"), workbook.renderSheetSvg(0, image));
            Files.write(outputDirectory.resolve("rendering.jpg"), Renderer.toJpeg(workbook, 0, image));
            Files.write(outputDirectory.resolve("rendering.tiff"), Renderer.toTiff(workbook, 0, image));
            workbook.saveHtml(outputDirectory.resolve("rendering.html"),
                    HtmlOptions.create().singleFile(true));
        }
    }
}

