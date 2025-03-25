package ru.japp.documents.excel.words.convert;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.FileOutputStream;

/**
 * Пример конвертации в формат RTF.
 */
public class DocxToRtf {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream("output.rtf")) {

            // Создание простого RTF
            doc.createParagraph().createRun().setText("Пример RTF");
            doc.write(out);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}