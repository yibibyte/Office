package ru.japp.documents.excel.words.convert;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * Конвертация файлов MS Word в HTML.
 */
public class DocxToHtml {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get("input.docx")));
             FileOutputStream out = new FileOutputStream("output.html")) {

            // Извлечение текста с базовым форматированием
            XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
            String html = "<html><body>"
                    + extractor.getText().replace("\n", "<br>")
                    + "</body></html>";

            out.write(html.getBytes());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}