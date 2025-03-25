package ru.japp.documents.excel;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;

public class StartExcel {
    public static void main(String[] args) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Данные");

            // Добавляем данные
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Месяц");
            header.createCell(1).setCellValue("Продажи");

            for (int i = 1; i <= 5; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue("Месяц " + i);
                row.createCell(1).setCellValue(i * 1000);
            }

            // Создаем диаграмму
            ExcelChartBuilder.addBarChart(sheet, "Продажи по месяцам");

            // Сохраняем файл
            try (FileOutputStream out = new FileOutputStream("Sales_Report.xlsx")) {
                workbook.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}