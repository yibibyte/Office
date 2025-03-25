package ru.japp.documents.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyleManager {
    private final XSSFWorkbook workbook;

    public ExcelStyleManager(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public CellStyle createHeaderStyle(IndexedColors color, short fontSize) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();

        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        font.setColor(color.getIndex());

        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);

        return style;
    }

    public CellStyle createDataStyle(String dataFormat) {
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(
                workbook.createDataFormat().getFormat(dataFormat)
        );
        return style;
    }

    public CellStyle createBordersStyle(IndexedColors borderColor) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setTopBorderColor(borderColor.getIndex());
        style.setBottomBorderColor(borderColor.getIndex());
        style.setLeftBorderColor(borderColor.getIndex());
        style.setRightBorderColor(borderColor.getIndex());
        return style;
    }
}