package ru.japp.documents.pdf;

import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;

import java.awt.*;
import java.io.IOException;
import java.util.List;

public class ReportGenerator {
    private final PdfBuilder pdfBuilder;
    private final PdfStyleConfig titleStyle;
    private final PdfStyleConfig bodyStyle;

    public ReportGenerator(String fontPath) throws IOException {
        this.pdfBuilder = new PdfBuilder();
        PDType0Font font = pdfBuilder.loadFont(fontPath);
        this.titleStyle = new PdfStyleConfig(font, 16, Color.BLACK);
        this.bodyStyle = new PdfStyleConfig(font, 12, new Color(50, 50, 50));
    }

    public void generateReport(String filePath, List<String> content) throws IOException {
        pdfBuilder.createPage(PDRectangle.A4);
        pdfBuilder.beginContent();

        // Заголовок
        pdfBuilder.setStyle(titleStyle);
        pdfBuilder.addText("Ежемесячный финансовый отчет", 50);
        pdfBuilder.addDivider();

        // Основной контент
        pdfBuilder.setStyle(bodyStyle);
        for (String item : content) {
            pdfBuilder.addText(item, 50);
        }

        // Логотип
        pdfBuilder.addImage("logo.png", 100);

        pdfBuilder.endContent();
        pdfBuilder.saveDocument(filePath);
    }
}