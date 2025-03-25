package ru.japp.documents.excel.convert;

import java.io.IOException;

/**
 * StartConvert - класс для запуска процесса конвертации файлов Excel в другие форматы.
 */
public class StartConvert {
    public static void main(String[] args) {
        try {
            new ExcelCsvConverter().convertExcelToCsv("input.xlsx", "output.csv");
            new ExcelPdfConverter().convertExcelToPdf("input.xlsx", "output.pdf");
            new ExcelJsonConverter().convertExcelToJson("input.xlsx", "output.json");
            new ExcelHtmlConverter().convertExcelToHtml("input.xlsx", "output.html");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}