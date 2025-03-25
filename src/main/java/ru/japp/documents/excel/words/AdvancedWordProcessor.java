package ru.japp.documents.excel.words;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.util.Units;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 *  Создание документа в формате .docx (MS Word) с использованием библиотеки Apache POI
 */
public class AdvancedWordProcessor {

    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument();
             FileOutputStream out = new FileOutputStream("AdvancedDocument.docx")) {

            // ==================== Основное содержимое ====================

            // Заголовок с стилями
            XWPFParagraph titleMain = doc.createParagraph();
            titleMain.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleMainRun = titleMain.createRun();
            titleMainRun.setText("Технический отчёт");
            titleMainRun.setBold(true);
            titleMainRun.setFontSize(20);
            titleMainRun.setColor("2F5496");

            // Заголовок 1-го уровня
            XWPFParagraph title = doc.createParagraph();
            title.setStyle("Heading1");
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Полное руководство по Apache POI");
            titleRun.setColor("2F5496");
            titleRun.setBold(true);
            titleRun.setFontSize(24);

            // ==================== Форматирование текста ====================
            XWPFParagraph textFormatting = doc.createParagraph();

            // Обычный текст
            XWPFRun normalRun = textFormatting.createRun();
            normalRun.setText("Стандартный текст. ");

            // Жирный + курсив
            XWPFRun styledRun = textFormatting.createRun();
            styledRun.setBold(true);
            styledRun.setItalic(true);
            styledRun.setText("Жирный и курсивный текст. ");

            // Подчеркнутый цветной текст
            XWPFRun coloredRun = textFormatting.createRun();
            coloredRun.setUnderline(UnderlinePatterns.SINGLE);
            coloredRun.setColor("FF0000");
            coloredRun.setText("Красный подчеркнутый текст.");

            // ==================== Таблицы ====================

            // Создание таблицы 3x3
            XWPFTable table = doc.createTable(3, 3);

            // Заголовок таблицы
            table.getRow(0).getCell(0).setText("№ ");
            table.getRow(0).getCell(1).setText("Наименование ");
            table.getRow(0).getCell(2).setText("Цена ");

            // Данные таблицы
            table.getRow(1).getCell(0).setText("1");
            table.getRow(1).getCell(1).setText("Ноутбук ");
            table.getRow(1).getCell(2).setText("1000 $ ");

            // ==================== Гиперссылки ====================
            XWPFParagraph linkPara = doc.createParagraph();
            XWPFHyperlinkRun linkRun = linkPara.createHyperlinkRun("https://poi.apache.org/components/");
            linkRun.setText("Пример ссылки на сайт Apache POI");
            linkRun.setUnderline(UnderlinePatterns.SINGLE);
            linkRun.setColor("0000FF");
//            linkRun.setColor("00CC00");
//            linkRun.setColor("FFAA00");
//            linkRun.setColor("FF0000");
//            linkRun.setColor("FFF500");
//            linkRun.setColor("9900ff");
//            linkRun.setColor("7945D6");

            // ==================== Изображения ====================
            if (Files.exists(Paths.get("logo.png"))) {
                try (InputStream is = Files.newInputStream(Paths.get("apache-original-wordmark.png"))) {
                    XWPFParagraph imagePara = doc.createParagraph();
                    XWPFRun imageRun = imagePara.createRun();
                    imageRun.addPicture(
                            is,
                            Document.PICTURE_TYPE_PNG,
                            "logo.png",
                            Units.toEMU(100),
                            Units.toEMU(75)
                    );
                }
            }

            // ==================== Колонтитулы ====================

            // Верхний колонтитул
            XWPFHeader header = doc.createHeader(HeaderFooterType.DEFAULT);
            XWPFParagraph headerPara = header.createParagraph();
            headerPara.setAlignment(ParagraphAlignment.RIGHT);
            headerPara.createRun().setText("Конфиденциально");

            // Нижний колонтитул с нумерацией
            XWPFFooter footer = doc.createFooter(HeaderFooterType.DEFAULT);
            XWPFParagraph footerPara = footer.createParagraph();
            footerPara.setAlignment(ParagraphAlignment.CENTER);
            footerPara.createRun().setText("Страница ");
            footerPara.getCTP().addNewFldSimple().setInstr("PAGE   \\* MERGEFORMAT");

            // ==================== Разрыв страницы ====================
            XWPFParagraph pageBreak = doc.createParagraph();
            pageBreak.setPageBreak(true);

            // ==================== Сохранение документа ====================
            doc.write(out);
            System.out.println("\n Документ успешно создан!");

        } catch (Exception e) {
            System.err.println("\n Ошибка при создании документа: " + e.getMessage());
            e.printStackTrace();
        }
    }
}