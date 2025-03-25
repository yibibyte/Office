package ru.japp.documents.excel.convert;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

/**
 * Создание Excel файла с помощью Apache POI.
 */
public class ExcelFileCreator {
    public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream("input.xlsx")) {

            // Создаем лист и стили
            Sheet sheet = workbook.createSheet("Employees");
            CreationHelper createHelper = workbook.getCreationHelper();

            // Стиль для дат
            CellStyle dateStyle = workbook.createCellStyle();
            dateStyle.setDataFormat(
                    createHelper.createDataFormat().getFormat("dd.MM.yyyy")
            );

            // Заголовки
            String[] headers = {"Имя", "Возраст", "Дата рождения", "Зарплата"};
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }

            // Данные
            Object[][] data = {
                    {"Иван Петров", 28, LocalDate.of(1995, 5, 15), 85000},
                    {"Мария Сидорова", 32, LocalDate.of(1991, 8, 23), 95000},
                    {"Алексей Иванов", 25, LocalDate.of(1998, 2, 10), 75000}
            };

            // Заполняем данные
            int rowNum = 1;
            for (Object[] rowData : data) {
                Row row = sheet.createRow(rowNum++);

                // Имя
                row.createCell(0).setCellValue((String) rowData[0]);

                // Возраст
                row.createCell(1).setCellValue((Integer) rowData[1]);

                // Дата рождения
                Cell dateCell = row.createCell(2);
                LocalDate localDate = (LocalDate) rowData[2];
                Date date = Date.from(
                        localDate.atStartOfDay(ZoneId.systemDefault()).toInstant()
                );
                dateCell.setCellValue(date);
                dateCell.setCellStyle(dateStyle);

                // Зарплата
                row.createCell(3).setCellValue((Integer) rowData[3]);
            }

            // Автонастройка ширины колонок
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(fos);
            System.out.println("Файл input.xlsx успешно создан!");

        } catch (Exception e) {
            System.err.println("Ошибка при создании файла: " + e.getMessage());
            e.printStackTrace();
        }
    }
}