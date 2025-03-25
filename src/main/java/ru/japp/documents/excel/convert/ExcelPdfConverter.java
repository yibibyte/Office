package ru.japp.documents.excel.convert;

import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import java.io.*;

/**
 * Класс для конвертации Excel файла в PDF.
 */
public class ExcelPdfConverter {

    public void convertExcelToPdf(String excelPath, String pdfPath) throws IOException {
        try (PDDocument pdfDoc = new PDDocument();
             Workbook workbook = new XSSFWorkbook(new File(excelPath));
             InputStream fontStream = getClass().getResourceAsStream("/fonts/arial.ttf")) {

            // Загрузка шрифта
            PDType0Font font = PDType0Font.load(pdfDoc, fontStream);

            Sheet sheet = workbook.getSheetAt(0);
            PDPage page = new PDPage();
            pdfDoc.addPage(page);

            try (PDPageContentStream contentStream = new PDPageContentStream(pdfDoc, page)) {
                contentStream.setFont(font, 12);
                contentStream.beginText();
                contentStream.newLineAtOffset(50, 700);

                for (Row row : sheet) {
                    String rowText = convertRowToText(row);
                    contentStream.showText(rowText); // Теперь кириллица работает!
                    contentStream.newLineAtOffset(0, -15);
                }

                contentStream.endText();
            }
            pdfDoc.save(pdfPath);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    private String convertRowToText(Row row) {
        StringBuilder sb = new StringBuilder();
        for (Cell cell : row) {
            sb.append(cell.toString()).append(" | ");
        }
        return !sb.isEmpty() ? sb.substring(0, sb.length() - 3) : "";
    }
}