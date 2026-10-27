package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.CellValue;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.Inflater;
import java.util.zip.ZipFile;

/**
 * Helpers for the issue tests: reading parts of a saved package and PDF content streams.
 *
 * <p>The tests run on the evaluation build, which adds an "Axyra Evaluation Copy" sheet to every saved
 * workbook. Sheets are therefore looked up by name, and worksheet XML is searched across every
 * worksheet part instead of a fixed {@code sheet1.xml}.
 */
final class IssueTestSupport {
    private IssueTestSupport() {
    }

    static CellValue n(double value) {
        return CellValue.number(value);
    }

    static CellValue t(String value) {
        return CellValue.text(value);
    }

    /** The text of one package part, or {@code null} when the part does not exist. */
    static String zipEntry(Path file, String name) throws IOException {
        try (ZipFile zip = new ZipFile(file.toFile())) {
            var entry = zip.getEntry(name);
            return entry == null ? null
                    : new String(zip.getInputStream(entry).readAllBytes(), StandardCharsets.UTF_8);
        }
    }

    static List<String> zipNames(Path file) throws IOException {
        try (ZipFile zip = new ZipFile(file.toFile())) {
            List<String> names = new ArrayList<>();
            zip.stream().forEach(entry -> names.add(entry.getName()));
            return names;
        }
    }

    /** Every worksheet part of the package, concatenated. */
    static String sheetsXml(Path file) throws IOException {
        StringBuilder xml = new StringBuilder();
        for (String name : zipNames(file)) {
            if (name.startsWith("xl/worksheets/sheet")) {
                xml.append(zipEntry(file, name));
            }
        }
        return xml.toString();
    }

    /** Every match of {@code regex} in {@code text}, in order. */
    static List<String> findAll(String text, String regex) {
        Matcher matcher = Pattern.compile(regex, Pattern.DOTALL).matcher(text);
        List<String> matches = new ArrayList<>();
        while (matcher.find()) {
            matches.add(matcher.group());
        }
        return matches;
    }

    static int count(String text, String part) {
        int count = 0;
        for (int i = text.indexOf(part); i >= 0; i = text.indexOf(part, i + part.length())) {
            count++;
        }
        return count;
    }

    /** The numeric value of a cell, or {@code NaN} when it does not hold a number. */
    static double number(CellValue value) {
        return value instanceof CellValue.Number number ? number.value() : Double.NaN;
    }

    /** Inflates every FlateDecode stream of a PDF and returns them in file order. */
    static List<String> pdfStreams(byte[] pdf) {
        String text = new String(pdf, StandardCharsets.ISO_8859_1);
        List<String> streams = new ArrayList<>();
        int i = 0;
        while ((i = text.indexOf("stream", i)) >= 0) {
            if (i >= 3 && text.startsWith("end", i - 3)) {
                i += 6;
                continue;
            }
            int start = i + 6;
            if (text.charAt(start) == '\r') {
                start++;
            }
            if (text.charAt(start) == '\n') {
                start++;
            }
            int end = text.indexOf("endstream", start);
            if (end < 0) {
                break;
            }
            Inflater inflater = new Inflater();
            inflater.setInput(pdf, start, end - start);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            byte[] buffer = new byte[8192];
            try {
                while (!inflater.finished()) {
                    int k = inflater.inflate(buffer);
                    if (k == 0 && (inflater.needsInput() || inflater.needsDictionary())) {
                        break;
                    }
                    out.write(buffer, 0, k);
                }
            } catch (Exception ignored) {
                // not a FlateDecode stream
            }
            streams.add(out.toString(StandardCharsets.ISO_8859_1));
            i = end + 9;
        }
        return streams;
    }
}
