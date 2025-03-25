package ru.japp.documents.excel.convert;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.*;
import java.util.*;

/**
 * Класс для конвертации между форматами Excel и JSON.
 */
public class ExcelJsonConverter {

    public void convertExcelToJson(String excelPath, String jsonPath) throws IOException {
        List<Map<String, String>> data = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new File(excelPath))) {
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Map<String, String> rowData = new HashMap<>();
                for (Cell cell : row) {
                    String header = headerRow.getCell(cell.getColumnIndex()).toString();
                    rowData.put(header, cell.toString());
                }
                data.add(rowData);
            }
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
        new ObjectMapper().writeValue(new File(jsonPath), data);
    }

    public void convertJsonToExcel(String jsonPath, String excelPath) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        List<Map<String, String>> data = mapper.readValue(
                new File(jsonPath),
                new com.fasterxml.jackson.core.type.TypeReference<List<Map<String, String>>>(){}
        );

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(excelPath)) {

            Sheet sheet = workbook.createSheet("Data");
            int rowNum = 0;
            for (Map<String, String> rowData : data) {
                Row row = sheet.createRow(rowNum++);
                int cellNum = 0;
                for (String value : rowData.values()) {
                    row.createCell(cellNum++).setCellValue(value);
                }
            }
            workbook.write(fos);
        }
    }
}