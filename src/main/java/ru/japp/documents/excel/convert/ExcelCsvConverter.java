package ru.japp.documents.excel.convert;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

/**
 *  Конвертер между файлами формата Excel и CSV.
 */

public class ExcelCsvConverter {

    public void convertExcelToCsv(String excelPath, String csvPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(new File(excelPath));
             FileWriter writer = new FileWriter(csvPath)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                String line = convertRowToCsv(row);
                writer.write(line + "\n");
            }
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public void convertCsvToExcel(String csvPath, String excelPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(excelPath)) {

            Sheet sheet = workbook.createSheet("Data");
            BufferedReader reader = new BufferedReader(new FileReader(csvPath));

            int rowNum = 0;
            String line;
            while ((line = reader.readLine()) != null) {
                Row row = sheet.createRow(rowNum++);
                String[] values = line.split(",");
                for (int i = 0; i < values.length; i++) {
                    row.createCell(i).setCellValue(values[i]);
                }
            }
            workbook.write(fos);
        }
    }

    private String convertRowToCsv(Row row) {
        StringBuilder sb = new StringBuilder();
        for (Cell cell : row) {
            sb.append(cell.toString()).append(",");
        }
        return sb.length() > 0 ? sb.deleteCharAt(sb.length() - 1).toString() : "";
    }
}