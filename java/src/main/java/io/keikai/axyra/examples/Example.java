package io.keikai.axyra.examples;

import java.nio.file.Path;

/** A runnable example associated with one Axyra Sheets documentation page. */
public interface Example {
    String name();

    String description();

    void run(Path outputDirectory) throws Exception;
}

