package io.keikai.axyra.examples.featuredemo;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;

/** Reads the input workbooks the feature demos start from, packaged under {@code /feature-demo/}. */
final class DemoBooks {
    private static final String FOLDER = "/feature-demo/";

    private DemoBooks() {
    }

    static byte[] read(String fileName) {
        try (InputStream in = DemoBooks.class.getResourceAsStream(FOLDER + fileName)) {
            if (in == null) {
                throw new IllegalStateException("book not found on the classpath: " + FOLDER + fileName);
            }
            return in.readAllBytes();
        } catch (IOException e) {
            throw new UncheckedIOException("cannot read " + FOLDER + fileName, e);
        }
    }
}
