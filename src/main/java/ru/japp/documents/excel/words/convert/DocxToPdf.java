package ru.japp.documents.excel.words.convert;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import java.io.File;
import java.io.FileOutputStream;

/**
 * Created by Dmitry on 17.07.2017.
 *  Конвертация файла в формате DOCX в PDF с помощью библиотеки docx4j
 */
public class DocxToPdf {
    public static void main(String[] args) {
        try {
            // Загрузка DOCX
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(
                    new File("E:\\Java\\Offece\\Documents\\input.docx")
            );

            // Конвертация в PDF
            Docx4J.toPDF(
                    wordMLPackage,
                    new FileOutputStream("output.pdf")
            );

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}