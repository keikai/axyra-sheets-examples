package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.examples.ExampleSupport;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.FormulaPolicy;
import io.keikai.axyra.sheets.io.SaveOptions;

import java.nio.file.Files;
import java.nio.file.Path;

public final class ImportExportExample implements Example {
    @Override
    public String name() {
        return "dev-ref/import-export";
    }

    @Override
    public String description() {
        return "Open from paths and bytes, then save XLSX, CSV, and JSON";
    }

    @Override
    public void run(Path outputDirectory) throws Exception {
        Path source = ExampleSupport.createSalesWorkbook(outputDirectory.resolve("import-source.xlsx"));
        byte[] bytes = Files.readAllBytes(source);
        try (Workbook workbook = Workbook.openBytes(bytes, "xlsx")) {
            workbook.recalculate();
            workbook.save(outputDirectory.resolve("values-only.xlsx"), SaveOptions.create()
                    .formulaPolicy(FormulaPolicy.VALUES_ONLY));
            workbook.save(outputDirectory.resolve("sales.csv"));
            Files.write(outputDirectory.resolve("sales.json"), workbook.saveBytes("json"));
        }
    }
}

