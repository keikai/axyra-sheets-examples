package io.keikai.axyra.examples;

import io.keikai.axyra.examples.devref.CellAndRangeExample;
import io.keikai.axyra.examples.devref.ChartsExample;
import io.keikai.axyra.examples.devref.ContentObjectsExample;
import io.keikai.axyra.examples.devref.FormulasExample;
import io.keikai.axyra.examples.devref.ImportExportExample;
import io.keikai.axyra.examples.devref.LicenseAndNativeLoaderExample;
import io.keikai.axyra.examples.devref.PivotTableExample;
import io.keikai.axyra.examples.devref.RenderingReferenceExample;
import io.keikai.axyra.examples.devref.SmartMarkersExample;
import io.keikai.axyra.examples.devref.StreamingExample;
import io.keikai.axyra.examples.devref.StylesExample;
import io.keikai.axyra.examples.devref.WorkbookAndSheetExample;
import io.keikai.axyra.examples.guides.CreateExcelFileExample;
import io.keikai.axyra.examples.guides.ExcelToPdfExample;
import io.keikai.axyra.examples.guides.GenerateReportExample;
import io.keikai.axyra.examples.guides.ReadExcelFileExample;
import io.keikai.axyra.examples.guides.WriteDataToExcelExample;
import io.keikai.axyra.examples.tutorial.QuickStartExample;
import io.keikai.axyra.examples.tutorial.ReadWriteExample;
import io.keikai.axyra.examples.tutorial.RenderExample;

import java.nio.file.Path;
import java.util.Comparator;
import java.util.List;

public final class ExampleRunner {
    private static final List<Example> EXAMPLES = List.of(
            new QuickStartExample(),
            new ReadWriteExample(),
            new RenderExample(),
            new CreateExcelFileExample(),
            new ReadExcelFileExample(),
            new ExcelToPdfExample(),
            new WriteDataToExcelExample(),
            new GenerateReportExample(),
            new WorkbookAndSheetExample(),
            new CellAndRangeExample(),
            new FormulasExample(),
            new StylesExample(),
            new ChartsExample(),
            new ContentObjectsExample(),
            new PivotTableExample(),
            new ImportExportExample(),
            new RenderingReferenceExample(),
            new StreamingExample(),
            new SmartMarkersExample(),
            new LicenseAndNativeLoaderExample());

    private ExampleRunner() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length == 0 || "list".equals(args[0])) {
            printExamples();
            return;
        }

        Path output = ExampleSupport.prepareOutputDirectory(
                args.length > 1 ? Path.of(args[1]) : Path.of("output"));
        if ("all".equals(args[0])) {
            for (Example example : EXAMPLES) run(example, output);
            return;
        }

        Example selected = EXAMPLES.stream()
                .filter(example -> example.name().equals(args[0]))
                .findFirst()
                .orElseThrow(() -> new IllegalArgumentException(
                        "Unknown example '" + args[0] + "'. Run without arguments to list examples."));
        run(selected, output);
    }

    public static List<Example> examples() {
        return EXAMPLES;
    }

    private static void run(Example example, Path output) throws Exception {
        System.out.println("\n== " + example.name() + " ==");
        example.run(output);
    }

    private static void printExamples() {
        System.out.println("Axyra Sheets documentation examples:\n");
        EXAMPLES.stream().sorted(Comparator.comparing(Example::name))
                .forEach(example -> System.out.printf("  %-38s %s%n",
                        example.name(), example.description()));
        System.out.println("\nRun one with: mvn exec:java -Dexec.args=\"<name> [output-directory]\"");
    }
}

