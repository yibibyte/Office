package ru.japp.documents.excel.words.convert;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @apiNote Convert docx to txt
 */
public class DocxToTxt {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument(
                Files.newInputStream(Paths.get("input.docx")));
             XWPFWordExtractor extractor = new XWPFWordExtractor(doc)) {

            Files.writeString(
                    Paths.get("output.txt"),
                    extractor.getText()
            );

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}