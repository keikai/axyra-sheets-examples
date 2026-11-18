import java.math.BigDecimal;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public final class SourceRecipes {
    public static List<WriteExcelFile.Sale> readSales(ResultSet rs)
            throws SQLException {
        List<WriteExcelFile.Sale> sales = new ArrayList<>();
        while (rs.next()) {
            BigDecimal amount = rs.getBigDecimal("amount");
            if (amount == null) {
                throw new IllegalArgumentException("amount is required");
            }
            sales.add(new WriteExcelFile.Sale(rs.getString("order_id"),
                    rs.getString("product"), amount.doubleValue()));
        }
        return sales;
    }
}
