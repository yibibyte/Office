package ru.japp.documents.pdf;

import java.io.IOException;
import java.util.List;

/**
 * Класс для запуска генерации PDF-отчета.
 */

public class StartPDF {
    public static void main(String[] args) {
        try {
            // Инициализация генератора
            ReportGenerator reportGenerator = new ReportGenerator("fonts/arial.ttf");
            // Подготовка данных
            List<String> reportData = List.of(
                    "Общий доход: 1 200 000 руб.",
                    "Расходы: 850 000 руб.",
                    "Чистая прибыль: 350 000 руб.",
                    "Топ товар: Ноутбуки (250 шт.)"
            );

            // Генерация отчета
            reportGenerator.generateReport("financial_report.pdf", reportData);

            System.out.println("PDF-отчет успешно создан!");
        } catch (IOException e) {
            System.err.println("Ошибка генерации отчета: " + e.getMessage());
            e.printStackTrace();
        }
    }
}