package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;

class RemainingExamplesIntegrationTest {
    @Test
    void runsEveryTutorialAndDeveloperReferenceExample(@TempDir Path output) throws Exception {
        List<Example> examples = ExampleRunner.examples().stream()
                .filter(example -> !example.name().startsWith("guides/"))
                .filter(example -> !example.name().equals("tutorial/quick-start"))
                .toList();

        assertFalse(examples.isEmpty());
        for (Example example : examples) {
            Path exampleOutput = Files.createDirectories(
                    output.resolve(example.name().replace('/', '-')));
            example.run(exampleOutput);
        }
    }
}
