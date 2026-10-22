package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Import Java objects straight onto a sheet.
 *
 * <p>{@code importData} reads a {@code List} of JavaBeans or records and writes
 * it as a block — no per-cell loop, no manual header row. {@link ImportOptions}
 * picks which properties go in, in what order, under what labels, and how the
 * header is styled.
 *
 * <p>A {@code Map} is a row too, keyed by column name: {@link #maps()} imports
 * records that have no class behind them with the same call.
 */
public final class ImportObjectsExample implements Example {
    @Override
    public String name() {
        return "feature-demo/import-objects";
    }

    @Override
    public String description() {
        return "Import a list of beans or maps onto a sheet with importData";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = beans()) {
            wb.save(outputDirectory.resolve("import-objects-beans.xlsx"));
        }
        try (Workbook wb = chosenColumns()) {
            wb.save(outputDirectory.resolve("import-objects-chosen-columns.xlsx"));
        }
        try (Workbook wb = maps()) {
            wb.save(outputDirectory.resolve("import-objects-maps.xlsx"));
        }
        try (Workbook wb = importAndTotal()) {
            wb.save(outputDirectory.resolve("import-objects-total.xlsx"));
        }
        System.out.println("Wrote four object-import workbooks to " + outputDirectory);
    }

    /** All properties, a styled header row, one call. */
    private static Workbook beans() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Trips");

        int rows = sheet.importData(0, 0, trips(), ImportOptions.builder()
                .headers(true)
                .columnLabels("City", "Nights", "Cost")
                .headerStyle(CellStyle.builder().bold()
                        .fill(Color.rgb(0xDD, 0xEB, 0xF7)).build())
                .build());

        sheet.cell(rows + 2, 0).setValue(CellValue.text(
                "importData wrote " + rows + " rows from a List<Trip>."));
        widen(sheet);
        return wb;
    }

    /** A subset, in an order of your choosing, under your own labels. */
    private static Workbook chosenColumns() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Costs");

        sheet.importData(0, 0, trips(), ImportOptions.builder()
                .columns("cost", "city")
                .headers(true)
                .columnLabels("Cost", "Destination")
                .headerStyle(CellStyle.builder().bold().build())
                .build());

        sheet.cell(trips().size() + 2, 0).setValue(CellValue.text(
                "columns(\"cost\", \"city\") reordered and dropped a property."));
        widen(sheet);
        return wb;
    }

    /** Keyed rows with no class behind them: each Map is a row, its keys the columns. */
    private static Workbook maps() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Dynamic");

        List<Map<String, Object>> rows = new ArrayList<>();
        rows.add(row("alpha", 12, "open"));
        rows.add(row("beta", 47, "closed"));
        rows.add(row("gamma", 8, "open"));

        // A LinkedHashMap fixes the column order; numbers stay numeric.
        sheet.importData(0, 0, rows, ImportOptions.builder()
                .headers(true)
                .headerStyle(CellStyle.builder().bold().build())
                .build());

        sheet.cell(rows.size() + 2, 0).setValue(CellValue.text(
                "Keys became the header row, in insertion order."));
        widen(sheet);
        return wb;
    }

    /** The imported block is ordinary cells, so formulas apply on top of it. */
    private static Workbook importAndTotal() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Trips");

        List<Trip> trips = trips();
        int rows = sheet.importData(0, 0, trips, ImportOptions.builder()
                .headers(true)
                .columnLabels("City", "Nights", "Cost")
                .headerStyle(CellStyle.builder().bold().build())
                .build());

        int totalRow = rows + 1;
        sheet.cell(totalRow, 0).setValue(CellValue.text("Total"));
        sheet.cell(totalRow, 1).setFormula("SUM(B2:B" + rows + ")");
        sheet.cell(totalRow, 2).setFormula("SUM(C2:C" + rows + ")");
        sheet.cell(totalRow + 1, 0).setValue(CellValue.text("Cost per night"));
        sheet.cell(totalRow + 1, 2).setFormula(
                "C" + (totalRow + 1) + "/B" + (totalRow + 1));

        sheet.range(totalRow, 0, totalRow, 2).setStyle(CellStyle.builder().bold().build());
        sheet.range(1, 2, totalRow + 1, 2).setStyle(
                CellStyle.builder().numberFormat("#,##0.00").build());
        widen(sheet);
        wb.recalculate();
        return wb;
    }

    private static List<Trip> trips() {
        List<Trip> trips = new ArrayList<>();
        trips.add(new Trip("Lisbon", 4, 820));
        trips.add(new Trip("Tallinn", 2, 410));
        trips.add(new Trip("Osaka", 7, 2140));
        trips.add(new Trip("Reykjavik", 3, 1290));
        return trips;
    }

    private static Map<String, Object> row(String key, int count, String state) {
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("key", key);
        map.put("count", count);
        map.put("state", state);
        return map;
    }

    private static void widen(Sheet sheet) {
        sheet.setColumnWidth(0, 18);
        sheet.setColumnWidth(1, 12);
        sheet.setColumnWidth(2, 14);
    }

    /** A plain JavaBean: properties are discovered from the getters. A record works identically. */
    public static final class Trip {
        private final String city;
        private final int nights;
        private final double cost;

        public Trip(String city, int nights, double cost) {
            this.city = city;
            this.nights = nights;
            this.cost = cost;
        }

        public String getCity() {
            return city;
        }

        public int getNights() {
            return nights;
        }

        public double getCost() {
            return cost;
        }
    }
}
