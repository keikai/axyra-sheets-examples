package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;

import java.util.HashSet;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

class ExampleCatalogTest {
    @Test
    void everyExampleHasAUniqueNameAndDescription() {
        var names = new HashSet<String>();
        for (Example example : ExampleRunner.examples()) {
            assertFalse(example.name().isBlank());
            assertFalse(example.description().isBlank());
            names.add(example.name());
        }
        assertEquals(ExampleRunner.examples().size(), names.size());
    }
}

