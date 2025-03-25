package ru.japp.documents.excel.words.convert;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * Конвертация файла .docx в .md
 */
public class DocxToMarkdown {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(Paths.get("input.docx")));
             XWPFWordExtractor extractor = new XWPFWordExtractor(doc)) {

            StringBuilder md = new StringBuilder();

            // Обработка параграфов
            doc.getParagraphs().forEach(p -> {
                String text = p.getText();
                if (!text.isEmpty()) {
                    // Заголовки (на основе стиля)
                    if (p.getStyle() != null && p.getStyle().startsWith("Heading")) {
                        int level = Integer.parseInt(p.getStyle().replace("Heading", ""));
                        md.append("#".repeat(level)).append(" ").append(text).append("\n\n");
                    } else {
                        // Обычный текст
                        md.append(text).append("\n\n");
                    }
                }
            });

            // Сохранение в Markdown
            Files.writeString(Paths.get("output.md"), md.toString());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}