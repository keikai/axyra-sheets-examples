package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;

import java.util.HashSet;
import java.util.TreeSet;

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

    @Test
    void everyFeatureDemoIsCoveredByAnIntegrationTest() {
        var registered = new TreeSet<String>();
        for (Example example : ExampleRunner.examples()) {
            if (example.name().startsWith("feature-demo/")) {
                registered.add(example.name());
            }
        }
        var tested = new TreeSet<String>();
        tested.addAll(FeatureDemoBasicsIntegrationTest.EXPECTED_OUTPUTS.keySet());
        tested.addAll(FeatureDemoDataAndOutputIntegrationTest.EXPECTED_OUTPUTS.keySet());
        assertEquals(20, registered.size());
        assertEquals(registered, tested);
    }
}
