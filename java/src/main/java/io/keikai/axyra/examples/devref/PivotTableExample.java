package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.PivotShowDataAsType;
import io.keikai.axyra.sheets.PivotStyle;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;

import java.nio.file.Path;

public final class PivotTableExample implements Example {
    @Override
    public String name() {
        return "dev-ref/pivot-table";
    }

    @Override
    public String description() {
        return "Create, configure, filter, and refresh a pivot table";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet data = workbook.createSheet("Data");
            data.range("A1:E5").setValues(new CellValue[][] {
                    {text("Region"), text("Product"), text("Quarter"), text("Channel"), text("Amount")},
                    {text("APAC"), text("Widget"), text("Q1"), text("Direct"), number(1200)},
                    {text("EMEA"), text("Widget"), text("Q1"), text("Partner"), number(980)},
                    {text("APAC"), text("Gadget"), text("Q2"), text("Direct"), number(1450)},
                    {text("Americas"), text("Gadget"), text("Q2"), text("Partner"), number(1720)}
            });
            Sheet summary = workbook.createSheet("Summary");
            summary.addPivotTable("SalesPivot", "A1", "Data!A1:E5");
            PivotTableView pivot = summary.pivotTable("SalesPivot");
            pivot.addRowField("Region");
            pivot.addColumnField("Quarter");
            pivot.addDataField("Amount", AggregateFunction.SUM);
            pivot.setItemFilter("Region", "APAC", "EMEA", "Americas");
            pivot.setShowDataAs(0, PivotShowDataAsType.PERCENT_OF_TOTAL);
            pivot.setStyle(PivotStyle.of("PivotStyleMedium9").withRowStripes());
            int written = pivot.refresh();
            System.out.println("Pivot cells written: " + written);
            workbook.save(outputDirectory.resolve("pivot-table.xlsx"));
        }
    }

    private static CellValue text(String value) {
        return CellValue.text(value);
    }

    private static CellValue number(double value) {
        return CellValue.number(value);
    }
}

