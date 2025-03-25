package ru.japp.documents.pdf;

import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import java.awt.*;
import java.io.*;

public class PdfBuilder {
    private final PDDocument document;
    private PDPage currentPage;
    private PDPageContentStream contentStream;
    private PdfStyleConfig currentStyle;
    private float currentY;

    public PdfBuilder() {
        this.document = new PDDocument();
    }

    // Инициализация новой страницы
    public void createPage(PDRectangle size) throws IOException {
        currentPage = new PDPage(size);
        document.addPage(currentPage);
        currentY = size.getHeight() - 50; // Стартовая позиция Y
    }

    // Загрузка шрифта
    public PDType0Font loadFont(String fontPath) throws IOException {
        // Используем ClassLoader для загрузки из ресурсов
        InputStream fontStream = getClass().getClassLoader().getResourceAsStream(fontPath);
        if (fontStream == null) {
            throw new FileNotFoundException("Font not found: " + fontPath);
        }
        return PDType0Font.load(document, fontStream);
    }

    // Начало работы с контентом
    public void beginContent() throws IOException {
        contentStream = new PDPageContentStream(document, currentPage);
    }

    // Установка стиля
    public void setStyle(PdfStyleConfig style) throws IOException {
        this.currentStyle = style;
        contentStream.setFont(style.getFont(), style.getFontSize());
        contentStream.setNonStrokingColor(style.getTextColor());
    }

    // Добавление текста с авто позиционированием
    public void addText(String text, float x) throws IOException {
        contentStream.beginText();
        contentStream.newLineAtOffset(x, currentY);
        contentStream.showText(text);
        contentStream.endText();
        currentY -= currentStyle.getFontSize() + 5; // Смещение для следующего элемента
    }

    // Добавление изображения
    public void addImage(String imagePath, float width) throws IOException {
        PDImageXObject image = PDImageXObject.createFromFile(imagePath, document);
        float aspectRatio = (float) image.getHeight() / image.getWidth();
        contentStream.drawImage(image, 50, currentY - (width * aspectRatio), width, width * aspectRatio);
        currentY -= (width * aspectRatio) + 20;
    }

    // Добавление разделительной линии
    public void addDivider() throws IOException {
        contentStream.setLineWidth(0.5f);
        contentStream.moveTo(50, currentY);
        contentStream.lineTo(currentPage.getMediaBox().getWidth() - 50, currentY);
        contentStream.stroke();
        currentY -= 15;
    }

    // Завершение страницы
    public void endContent() throws IOException {
        contentStream.close();
    }

    // Сохранение документа
    public void saveDocument(String filePath) throws IOException {
        document.save(filePath);
        document.close();
    }
}