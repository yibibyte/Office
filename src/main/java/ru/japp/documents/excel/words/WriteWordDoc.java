package ru.japp.documents.excel.words;

import org.apache.poi.xwpf.usermodel.*;
import java.io.FileOutputStream;

/**
 * Запись данных в Word-файл
 */
public class WriteWordDoc {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream("output.docx")) {

            // Создание заголовка
            XWPFParagraph title = doc.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Отчёт о продажах");
            titleRun.setBold(true);
            titleRun.setFontSize(18);

            // Добавление текста
            XWPFParagraph paragraph = doc.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("Дата: 2024-01-01");
            run.addBreak();
            run.setText("Продажи за январь: 1 000 000 руб.");

            // Создание таблицы
            XWPFTable table = doc.createTable(3, 2); // 3 строки, 2 колонки
            table.getRow(0).getCell(0).setText("Товар");
            table.getRow(0).getCell(1).setText("Количество");
            table.getRow(1).getCell(0).setText("Ноутбуки");
            table.getRow(1).getCell(1).setText("150");
            table.getRow(2).getCell(0).setText("Смартфоны");
            table.getRow(2).getCell(1).setText("300");

            // Сохранение документа
            doc.write(out);
            System.out.println("Документ успешно создан!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}