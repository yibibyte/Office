package ru.japp.documents.excel;

import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelChartBuilder {
    public static void addBarChart(XSSFSheet sheet, String title) {
        // Создаем объект для рисования
        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 5, 0, 15, 20);

        // Создаем диаграмму
        XSSFChart chart = drawing.createChart(anchor);
        chart.setTitleText(title);

        // Настраиваем оси
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);

        // Источники данных
        XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(
                sheet, new CellRangeAddress(1, 5, 0, 0)
        );
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(
                sheet, new CellRangeAddress(1, 5, 1, 1)
        );

        // Создаем данные для диаграммы
        XDDFBarChartData data = (XDDFBarChartData) chart.createData(
                ChartTypes.BAR, bottomAxis, leftAxis
        );
        XDDFBarChartData.Series series = (XDDFBarChartData.Series) data.addSeries(categories, values);
        series.setTitle("Продажи", null);

        // Применяем данные к диаграмме
        chart.plot(data);
    }
}