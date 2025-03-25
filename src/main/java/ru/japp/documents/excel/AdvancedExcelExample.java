package ru.japp.documents.excel;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.io.FileOutputStream;
import java.time.LocalDate;

public class AdvancedExcelExample {
    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook()) {
            // ==================== 1. Создание листа ====================
            Sheet sheet = workbook.createSheet("Отчёт");

            // ==================== 2. Форматирование ячеек ====================
            // Стиль для заголовков
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.BLUE.getIndex());
            headerStyle.setFont(headerFont);

            // Стиль для дат
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(workbook.createDataFormat().getFormat("dd.MM.yyyy"));

            // ==================== 3. Заголовки таблицы ====================
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Дата", "Категория", "Сумма", "Формула"};

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // ==================== 4. Добавление данных ====================
            Object[][] data = {
                    {LocalDate.of(2024, 1, 5), "Продукты", 2500},
                    {LocalDate.of(2024, 1, 6), "Транспорт", 1200},
                    {LocalDate.of(2024, 1, 7), "Развлечения", 3500}
            };

            int rowNum = 1;
            for (Object[] rowData : data) {
                Row row = sheet.createRow(rowNum++);

                // Дата
                Cell dateCell = row.createCell(0);
                dateCell.setCellValue((LocalDate) rowData[0]);
                dateCell.setCellStyle(dateStyle);

                // Категория
                row.createCell(1).setCellValue((String) rowData[1]);

                // Сумма
                row.createCell(2).setCellValue((Integer) rowData[2]);
            }

            // ==================== 5. Формулы ====================
            Row totalRow = sheet.createRow(rowNum);
            totalRow.createCell(1).setCellValue("Итого:");
            totalRow.createCell(2).setCellFormula("SUM(C2:C4)");

            // ==================== 6. Гиперссылка ====================
            Row linkRow = sheet.createRow(rowNum + 1);
            Cell linkCell = linkRow.createCell(0);
            linkCell.setCellValue("Ссылка на сайт");
            Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress("https://example.com");
            linkCell.setHyperlink(link);

            // ==================== 7. Настройка ширины колонок ====================
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // ==================== 8. Сохранение файла ====================
            try (FileOutputStream out = new FileOutputStream("Financial_Report.xlsx")) {
                workbook.write(out);
                System.out.println("Файл успешно создан!");
            }

        } catch (Exception e) {
            System.err.println("Ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }
}