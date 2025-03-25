package ru.japp.documents.excel.words;

import org.apache.poi.xwpf.usermodel.*;
import java.io.FileOutputStream;

/**
 * Пример создания документа Word с помощью Apache POI.
 */
public class CreateWord {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument()) {
            // Заголовок
            XWPFParagraph title = doc.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Отчёт");
            titleRun.setBold(true);
            titleRun.setFontSize(16);

            // Текст
            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("Пример текста в документе Word.");
            run.addBreak(); // Перенос строки

            // Таблица
            XWPFTable table = doc.createTable(2, 2);
            table.getRow(0).getCell(0).setText("Колонка 1");
            table.getRow(0).getCell(1).setText("Колонка 2");
            table.getRow(1).getCell(0).setText("Данные 1");
            table.getRow(1).getCell(1).setText("Данные 2");

            // Сохранение
            try (FileOutputStream out = new FileOutputStream("report.docx")) {
                doc.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}