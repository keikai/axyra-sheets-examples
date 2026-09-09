package io.keikai.axyra.examples.guides;

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

public final class WriteDataToExcelExample implements Example {
    public record Sale(String region, LocalDate orderDate, double amount) {
    }

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
        List<Sale> sales = List.of(
                new Sale("APAC", LocalDate.of(2026, 9, 1), 1450),
                new Sale("EMEA", LocalDate.of(2026, 9, 2), 1100));

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Q3 Sales");
            CellStyle header = CellStyle.builder().bold().build();
            sheet.importData(0, 0, sales, ImportOptions.builder()
                    .columns("region", "orderDate", "amount")
                    .headers(true)
                    .columnLabels("Region", "Order date", "Amount")
                    .headerStyle(header)
                    .build());
            sheet.range("B2:B3").setStyle(CellStyle.builder()
                    .numberFormat("yyyy-mm-dd").build());

            JsonNode rows = new ObjectMapper().readTree("""
                    [{"region":"Americas","orderDate":"2026-09-03","amount":1720.0}]
                    """);
            CellValue[][] jsonValues = new CellValue[rows.size()][3];
            for (int row = 0; row < rows.size(); row++) {
                JsonNode value = rows.get(row);
                jsonValues[row] = new CellValue[] {
                        CellValue.text(value.get("region").asText()),
                        CellValue.text(value.get("orderDate").asText()),
                        CellValue.number(value.get("amount").asDouble())
                };
            }
            sheet.range(3, 0, 3 + rows.size() - 1, 2).setValues(jsonValues);

            Sheet metrics = workbook.createSheet("Numeric Matrix");
            metrics.range(0, 0, 1, 2).setNumbers(new double[] {
                    12, 1450, 17,
                    3, 1100, 9
            });
            workbook.save(outputDirectory.resolve("data-export.xlsx"));
        }

        writeStreamingExport(outputDirectory.resolve("streaming-export.xlsx"), sales);
    }

    private static void writeStreamingExport(Path output, List<Sale> sales) {
        try (StreamWorkbook workbook = StreamWorkbook.create(output)) {
            int header = workbook.addStyle(StreamStyle.defaults().withBold());
            int date = workbook.addStyle(
                    StreamStyle.defaults().withNumberFormat("yyyy-mm-dd"));
            int money = workbook.addStyle(
                    StreamStyle.defaults().withNumberFormat("#,##0.00"));

            workbook.startSheet("Sales");
            workbook.row(0)
                    .addText("Region", header)
                    .addText("Order date", header)
                    .addText("Amount", header)
                    .commit();

            for (int row = 0; row < sales.size(); row++) {
                Sale sale = sales.get(row);
                workbook.row(row + 1)
                        .addText(sale.region())
                        .addDate(sale.orderDate(), date)
                        .addNumber(sale.amount(), money)
                        .commit();
            }
        }
    }
}
