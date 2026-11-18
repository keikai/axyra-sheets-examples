package io.keikai.axyra.examples.guides;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.io.StreamStyle;
import io.keikai.axyra.sheets.io.StreamWorkbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;

/**
 * The code from "How to Write Data to Excel in Java". The guide maps a JDBC
 * {@code ResultSet} to records; this example starts from the records so that it
 * runs without a database.
 */
public final class WriteDataToExcelExample implements Example {
    public record Sale(String region, String product, LocalDate orderDate, double amount) {
    }

    private static final String JSON = """
            [{"region":"Americas","product":"Widget","amount":1720.0},
             {"region":"APAC","amount":null}]
            """;

    @Override
    public String name() {
        return "guides/write-data-to-excel";
    }

    @Override
    public String description() {
        return "Import Java records and JSON, and write a streaming Excel export";
    }

    @Override
    public void run(Path outputDirectory) throws Exception {
        List<Sale> rows = List.of(
                new Sale("APAC", "Widget", LocalDate.of(2026, 9, 1), 1450),
                new Sale("EMEA", "Gadget", LocalDate.of(2026, 9, 2), 1100));

        try (Workbook wb = Workbook.create()) {
            // Records: pinned columns, header row, then format the date and money columns.
            Sheet sheet = wb.createSheet("Q3 Sales");
            CellStyle headerStyle = CellStyle.builder().bold().build();
            int written = sheet.importData(0, 0, rows, ImportOptions.builder()
                    .columns("region", "product", "orderDate", "amount")
                    .headers(true)
                    .columnLabels("Region", "Product", "Order Date", "Amount")
                    .headerStyle(headerStyle)
                    .build());
            System.out.println(written + " rows including the header");

            CellStyle dateStyle = CellStyle.builder().numberFormat("yyyy-mm-dd").build();
            CellStyle money = CellStyle.builder().numberFormat("#,##0.00").build();
            int lastRow = rows.size();                           // header is row 0
            sheet.range(1, 2, lastRow, 2).setStyle(dateStyle);   // column C
            sheet.range(1, 3, lastRow, 3).setStyle(money);       // column D

            // JSON parsed into maps, imported directly.
            ObjectMapper mapper = new ObjectMapper();
            List<Map<String, Object>> maps = mapper.readValue(JSON,
                    new TypeReference<List<Map<String, Object>>>() {});
            Sheet fromMaps = wb.createSheet("JSON (maps)");
            fromMaps.importData(0, 0, maps, ImportOptions.builder()
                    .columns("region", "product", "amount")
                    .columnLabels("Region", "Product", "Amount")
                    .build());

            // JSON with full control over each cell's type.
            Sheet dynamic = wb.createSheet("JSON (dynamic)");
            JsonNode array = mapper.readTree(JSON);
            String[] columns = { "region", "product", "amount" };

            CellValue[][] header = new CellValue[1][columns.length];
            for (int c = 0; c < columns.length; c++) {
                header[0][c] = CellValue.text(columns[c]);
            }
            dynamic.range(0, 0, 0, columns.length - 1).setValues(header);

            CellValue[][] body = new CellValue[array.size()][columns.length];
            for (int r = 0; r < array.size(); r++) {
                JsonNode node = array.get(r);
                for (int c = 0; c < columns.length; c++) {
                    JsonNode field = node.get(columns[c]);
                    body[r][c] = field == null || field.isNull() ? CellValue.blank()
                            : field.isNumber() ? CellValue.number(field.asDouble())
                            : field.isBoolean() ? CellValue.bool(field.asBoolean())
                            : CellValue.text(field.asText());
                }
            }
            dynamic.range(1, 0, array.size(), columns.length - 1).setValues(body);

            // The fast path for numbers: a flat, row-major double[].
            int rowCount = 1_000, colCount = 12;
            double[] flat = new double[rowCount * colCount];
            for (int i = 0; i < flat.length; i++) {
                flat[i] = i;                                     // flat[r * colCount + c]
            }
            Sheet matrix = wb.createSheet("Numeric Matrix");
            matrix.range(1, 0, rowCount, colCount - 1).setNumbers(flat);

            wb.save(outputDirectory.resolve("data-export.xlsx"));
        }

        writeStreamingExport(outputDirectory.resolve("streaming-export.xlsx"), rows);
    }

    /** The streaming writer from the guide, fed from records instead of a ResultSet. */
    private static void writeStreamingExport(Path output, List<Sale> sales) {
        try (StreamWorkbook out = StreamWorkbook.create(output)) {
            int header = out.addStyle(StreamStyle.defaults().withBold());
            int money  = out.addStyle(StreamStyle.defaults().withNumberFormat("#,##0.00"));
            int date   = out.addStyle(StreamStyle.defaults().withNumberFormat("yyyy-mm-dd"));

            out.startSheet("Sales");

            out.row(0).addText("Region", header)
                      .addText("Order Date", header)
                      .addText("Amount", header)
                      .commit();

            int r = 1;
            double total = 0;
            for (Sale sale : sales) {
                out.row(r++)
                   .addText(sale.region())
                   .addDate(sale.orderDate(), date)
                   .addNumber(sale.amount(), money)
                   .commit();
                total += sale.amount();
            }
            // No recalculation in a stream: supply the cached value with the formula.
            out.row(r).setText(0, "Total", header)
                      .setFormula(2, "SUM(C2:C" + r + ")", total, money)
                      .commit();
        }
    }
}
