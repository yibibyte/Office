package ru.japp.documents.excel.convert;

import java.io.IOException;

public class MainConvert {
    public static void main(String[] args) {
        try {
            // Excel → CSV
            new ExcelCsvConverter().convertExcelToCsv("input.xlsx", "output.csv");

            // CSV → Excel
            new ExcelCsvConverter().convertCsvToExcel("input.csv", "output.xlsx");

            // Excel → PDF
            new ExcelPdfConverter().convertExcelToPdf("input.xlsx", "output.pdf");

            // Excel → JSON
            new ExcelJsonConverter().convertExcelToJson("input.xlsx", "output.json");

            // JSON → Excel
            new ExcelJsonConverter().convertJsonToExcel("input.json", "output.xlsx");

            // Excel → HTML
            new ExcelHtmlConverter().convertExcelToHtml("input.xlsx", "output.html");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}