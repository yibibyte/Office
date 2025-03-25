package ru.japp.documents.pdf;

import org.apache.pdfbox.pdmodel.font.PDType0Font;
import java.awt.*;
import java.io.IOException;

public class PdfStyleConfig {
    private PDType0Font font;
    private float fontSize;
    private Color textColor;

    public PdfStyleConfig(PDType0Font font, float fontSize, Color textColor) {
        this.font = font;
        this.fontSize = fontSize;
        this.textColor = textColor;
    }

    // Геттеры
    public PDType0Font getFont() { return font; }
    public float getFontSize() { return fontSize; }
    public Color getTextColor() { return textColor; }
}