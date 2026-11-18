import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import java.nio.file.Path;
import java.util.List;

public final class WriteJsonSalesReport {
    public static void main(String[] args) throws Exception {
        String json = """
            [{"region":"East","product":"Laptop",
              "orderDate":"2026-10-27","amount":1250.5}]
            """;
        ObjectMapper mapper = new ObjectMapper().registerModule(new JavaTimeModule());
        List<DatabaseSalesReport.Sale> rows = mapper.readValue(json,
                new TypeReference<List<DatabaseSalesReport.Sale>>() {});
        DatabaseSalesReport.writeReport(rows, Path.of("q3-json-sales.xlsx"));
    }
}
