package io.keikai.axyra.examples;

import io.keikai.axyra.examples.tutorial.QuickStartExample;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfSystemProperty;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertTrue;

class QuickStartIntegrationTest {
    @Test
    @EnabledIfSystemProperty(named = "axyra.native.path", matches = ".+")
    void writesARealWorkbook(@TempDir Path output) throws Exception {
        new QuickStartExample().run(output);
        assertTrue(Files.size(output.resolve("hello.xlsx")) > 0);
    }
}

