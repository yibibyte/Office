package ru.japp.documents.excel.words;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;

/**
 * Чтение текста из Word-файла
 */

public class ReadWordDoc {
    public static void main(String[] args) {
        File inputFile = new File("input.docx");

        // Проверка существования файла
        if (!inputFile.exists()) {
            System.out.println("Ошибка: Файл input.docx не найден!");
            return;
        }

        // Проверка, что это файл, а не директория
        if (!inputFile.isFile()) {
            System.out.println("Ошибка: input.docx не является файлом!");
            return;
        }

        // Проверка возможности чтения
        if (!inputFile.canRead()) {
            System.out.println("Ошибка: Нет прав на чтение файла!");
            return;
        }

        try (FileInputStream fis = new FileInputStream(inputFile);
             XWPFDocument doc = new XWPFDocument(fis)) {

            System.out.println("\n=== Текст из документа ===");

            // Чтение всех параграфов
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                String text = paragraph.getText();
                if (!text.isEmpty()) {
                    System.out.println(text);
                }
            }

            // Чтение текста из таблиц
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        System.out.print(cell.getText() + "\t");
                    }
                    System.out.println();
                }
            }

        } catch (Exception e) {
            System.out.println("Ошибка при чтении файла: " + e.getMessage());
            e.printStackTrace();
        }
    }
}