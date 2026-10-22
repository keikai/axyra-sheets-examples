package io.keikai.axyra.examples;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Runs a group of feature-demo examples and checks every file each one writes.
 *
 * <p>The demos are split into groups, one test class each, because the evaluation
 * build allows 100 opens, saves and renders per process and Surefire forks one JVM
 * per test class.
 */
final class FeatureDemoTestSupport {
    private static final byte[] PDF_MAGIC = "%PDF".getBytes(StandardCharsets.US_ASCII);
    private static final byte[] PNG_MAGIC = {(byte) 0x89, 'P', 'N', 'G'};
    private static final byte[] ZIP_MAGIC = {'P', 'K', 3, 4};

    private FeatureDemoTestSupport() {
    }

    static void runAndCheck(Map<String, List<String>> expectedOutputs, Path output) throws Exception {
        for (Map.Entry<String, List<String>> entry : expectedOutputs.entrySet()) {
            Example example = ExampleRunner.examples().stream()
                    .filter(candidate -> candidate.name().equals(entry.getKey()))
                    .findFirst()
                    .orElseThrow(() -> new AssertionError("no example named " + entry.getKey()));
            Path exampleOutput = Files.createDirectories(
                    output.resolve(example.name().replace('/', '-')));
            example.run(exampleOutput);
            for (String fileName : entry.getValue()) {
                checkFile(example.name(), exampleOutput.resolve(fileName));
            }
        }
    }

    /** The file exists and starts the way its format requires. */
    private static void checkFile(String exampleName, Path file) throws Exception {
        assertTrue(Files.isRegularFile(file), () -> exampleName + " did not produce " + file.getFileName());
        byte[] bytes = Files.readAllBytes(file);
        assertTrue(bytes.length > 0, () -> exampleName + " wrote an empty " + file.getFileName());
        String name = file.getFileName().toString();
        if (name.endsWith(".xlsx")) {
            assertArrayEquals(ZIP_MAGIC, Arrays.copyOf(bytes, 4), name + " is not a zip package");
        } else if (name.endsWith(".pdf")) {
            assertArrayEquals(PDF_MAGIC, Arrays.copyOf(bytes, 4), name + " is not a PDF");
        } else if (name.endsWith(".png")) {
            assertArrayEquals(PNG_MAGIC, Arrays.copyOf(bytes, 4), name + " is not a PNG");
        } else if (name.endsWith(".svg")) {
            assertTrue(new String(bytes, StandardCharsets.UTF_8).contains("<svg"), name + " is not an SVG");
        }
    }
}
