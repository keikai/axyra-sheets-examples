package io.keikai.axyra.examples.devref;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ChartView;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

public final class ChartsExample implements Example {
    @Override
    public String name() {
        return "dev-ref/charts";
    }

    @Override
    public String description() {
        return "Build an embedded chart and a chart sheet";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.createSheet("Sales");
            sheet.range("A1:C4").setValues(new CellValue[][] {
                    {CellValue.text("Region"), CellValue.text("Q1"), CellValue.text("Q2")},
                    {CellValue.text("APAC"), CellValue.number(1200), CellValue.number(1450)},
                    {CellValue.text("EMEA"), CellValue.number(980), CellValue.number(1100)},
                    {CellValue.text("Americas"), CellValue.number(1610), CellValue.number(1720)}
            });
            Chart chart = Chart.column()
                    .title("Revenue by Region")
                    .series("Q1", "Sales!$B$2:$B$4", "Sales!$A$2:$A$4")
                    .series("Q2", "Sales!$C$2:$C$4", "Sales!$A$2:$A$4")
                    .legend(LegendPosition.BOTTOM)
                    .showDataLabels(true)
                    .chartAreaFill(Color.parse("#FFFFFF"))
                    .anchor(1, 5, 16, 12)
                    .build();
            sheet.addChart(chart);

            int chartSheet = workbook.addChartSheet("Revenue", chart);
            ChartView view = workbook.chartSheetChart(chartSheet);
            System.out.println("Chart sheet: " + workbook.chartSheetName(chartSheet)
                    + "; title=" + view.title());
            workbook.save(outputDirectory.resolve("charts.xlsx"));
        }
    }
}

