package ru.japp.documents.excel.words;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;

/**
 * Создание документа с помощью Apache POI
 */
public class CreateAdvancedWord {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream("advanced.docx")) {

            // Заголовок с стилями
            XWPFParagraph title = doc.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Технический отчёт");
            titleRun.setBold(true);
            titleRun.setFontSize(20);
            titleRun.setColor("2F5496");

            // Таблица с объединенными ячейками
            XWPFTable table = doc.createTable(2, 2);
            table.getRow(0).getCell(0).setText("Данные 1");
            table.getRow(0).getCell(1).setText("Данные 2");

            // Сохранение
            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}