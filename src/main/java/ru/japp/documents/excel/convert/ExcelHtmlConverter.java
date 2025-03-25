package ru.japp.documents.excel.convert;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.time.format.DateTimeFormatter;
import java.time.ZoneId;
import java.util.Date;
import java.util.Locale;

/**
 * Класс для конвертации Excel-файла в HTML.
 */
public class ExcelHtmlConverter {

    public void convertExcelToHtml(String excelPath, String htmlPath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(new File(excelPath));
             FileOutputStream fos = new FileOutputStream(htmlPath)) {

            Sheet sheet = workbook.getSheetAt(0);
            StringBuilder html = new StringBuilder();

            // Заголовок HTML с указанием кодировки
            html.append("<!DOCTYPE html>\n")
                    .append("<html>\n<head>\n")
                    .append("  <meta charset='UTF-8'>\n")
                    .append("  <title>Excel to HTML</title>\n")
                    .append("  <style>\n")
                    .append("    table { border-collapse: collapse; width: 100%; }\n")
                    .append("    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }\n")
                    .append("    th { background-color: #f2f2f2; }\n")
                    .append("  </style>\n")
                    .append("</head>\n<body>\n");

            // Начало таблицы
            html.append("<table>\n");

            // Форматтер для дат
            DateTimeFormatter dateFormatter = DateTimeFormatter
                    .ofPattern("dd.MM.yyyy")
                    .withLocale(new Locale("ru"));

            // Обработка строк
            boolean isFirstRow = true;
            for (Row row : sheet) {
                html.append("  <tr>\n");
                for (Cell cell : row) {
                    String cellValue = formatCellValue(cell, dateFormatter);
                    if (isFirstRow) {
                        html.append("    <th>").append(cellValue).append("</th>\n");
                    } else {
                        html.append("    <td>").append(cellValue).append("</td>\n");
                    }
                }
                html.append("  </tr>\n");
                isFirstRow = false;
            }

            // Завершение HTML-документа
            html.append("</table>\n")
                    .append("</body>\n</html>");

            // Запись с явным указанием кодировки
            fos.write(html.toString().getBytes("UTF-8"));
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    private String formatCellValue(Cell cell, DateTimeFormatter dateFormatter) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // Форматирование даты
                    Date date = cell.getDateCellValue();
                    return date.toInstant()
                            .atZone(ZoneId.systemDefault())
                            .toLocalDate()
                            .format(dateFormatter);
                } else {
                    // Форматирование чисел
                    double num = cell.getNumericCellValue();
                    if (num == (int) num) {
                        return String.valueOf((int) num);
                    } else {
                        return String.format(Locale.US, "%.2f", num);
                    }
                }
            case STRING:
                return cell.getStringCellValue();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}