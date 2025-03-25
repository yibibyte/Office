package ru.japp.documents.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xddf.usermodel.*;

import java.io.FileOutputStream;

public class ExcelConditionalFormatting {
    private final Sheet sheet;
    private final XSSFWorkbook workbook;

    public ExcelConditionalFormatting(XSSFWorkbook workbook, Sheet sheet) {
        this.workbook = workbook;
        this.sheet = sheet;
    }

    public void addColorScaleFormatting(String range, int minColorIndex, int maxColorIndex) {
        // Создаем объект для условного форматирования
        SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();

        // Создаем правило с цветовой шкалой
        ConditionalFormattingRule rule = formatting.createConditionalFormattingColorScaleRule();

        // Настраиваем цвета через прямое указание индексов
        Color[] colors = {
                new XSSFColor(IndexedColors.fromInt(minColorIndex), workbook.getStylesSource().getIndexedColors()),
                new XSSFColor(IndexedColors.fromInt(maxColorIndex), workbook.getStylesSource().getIndexedColors())
        };

        // Для POI 5.2.5+ используем reflection для установки цветов
        try {
            ColorScaleFormatting colorFormatting = rule.getColorScaleFormatting();
            java.lang.reflect.Method setColors = colorFormatting.getClass()
                    .getMethod("setColors", Color[].class);
            setColors.invoke(colorFormatting, (Object) colors);
        } catch (Exception e) {
            throw new RuntimeException("Failed to set colors", e);
        }

        // Применяем к диапазону
        formatting.addConditionalFormatting(
                new CellRangeAddress[]{CellRangeAddress.valueOf(range)},
                rule
        );
    }

    // Пример использования
    public static void main(String[] args) {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");

            // Добавляем тестовые данные
            for (int i = 0; i < 10; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue(i * 10);
            }

            // Создаем форматирование
            ExcelConditionalFormatting cf = new ExcelConditionalFormatting(workbook, sheet);
            cf.addColorScaleFormatting("A1:A10",
                    IndexedColors.RED.getIndex(),
                    IndexedColors.GREEN.getIndex()
            );

            // Сохраняем файл
            try (FileOutputStream out = new FileOutputStream("ColorScaleExample.xlsx")) {
                workbook.write(out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}