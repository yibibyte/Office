package ru.japp.documents.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Чтение и запись Excel-файла
 */
public class CreateAndReadExcel {
    public static void main(String[] args) {
        // Запись в Excel
        try (Workbook workbook = new XSSFWorkbook()) { // .xlsx
            Sheet sheet = workbook.createSheet("Данные");

            // Создание строки
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Имя");
            row.createCell(1).setCellValue("Возраст");

            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("Анна");
            dataRow.createCell(1).setCellValue(28);

            try (FileOutputStream out = new FileOutputStream("data.xlsx")) {
                workbook.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Чтение из Excel
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream("data.xlsx"))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell + "\t");
                }
                System.out.println();
            }
        } catch (Exception e) {
            System.out.println("Ошибка чтения файла: " + e.getCause() + " " + e.getMessage());
        }
    }
}