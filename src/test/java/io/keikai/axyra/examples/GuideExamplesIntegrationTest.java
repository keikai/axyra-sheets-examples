package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class GuideExamplesIntegrationTest {
    private static final List<String> EXPECTED_GUIDES = List.of(
            "guides/convert-excel-to-pdf",
            "guides/create-excel-file",
            "guides/generate-report",
            "guides/read-excel-file",
            "guides/write-data-to-excel");
    private static final Map<String, List<String>> EXPECTED_OUTPUTS = Map.of(
            "guides/convert-excel-to-pdf", List.of("pdf-source.xlsx", "report.pdf"),
            "guides/create-excel-file", List.of("q3-sales.xlsx"),
            "guides/generate-report", List.of(
                    "sales-template.xlsx", "sales-Q3-2026.xlsx", "sales-Q3-2026.pdf"),
            "guides/read-excel-file", List.of("read-source.xlsx"),
            "guides/write-data-to-excel", List.of(
                    "data-export.xlsx", "streaming-export.xlsx"));

    @Test
    @EnabledIfSystemProperty(named = "axyra.native.path", matches = ".+")
    void runsEveryGuideExample(@TempDir Path output) throws Exception {
        List<Example> guides = ExampleRunner.examples().stream()
                .filter(example -> example.name().startsWith("guides/"))
                .toList();

        assertEquals(EXPECTED_GUIDES,
                guides.stream().map(Example::name).sorted().toList());

        for (Example guide : guides) {
            Path guideOutput = Files.createDirectories(
                    output.resolve(guide.name().substring("guides/".length())));
            guide.run(guideOutput);
            for (String fileName : EXPECTED_OUTPUTS.get(guide.name())) {
                Path file = guideOutput.resolve(fileName);
                assertTrue(isNonEmptyFile(file),
                        () -> guide.name() + " did not produce " + fileName);
            }
        }
    }

    private static boolean isNonEmptyFile(Path path) {
        try {
            return Files.isRegularFile(path) && Files.size(path) > 0;
        } catch (Exception ignored) {
            return false;
        }
    }
}
