package com.example.reportgenerator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Service
public class ReportService {

    // Генерация обычного текстового отчета
    public String generateReport(String reportType) {
        return "Отчет: " + reportType;
    }

    // Генерация отчета в формате .docx с таблицей
    public byte[] generateDocxReport(String reportType) throws IOException {
        // Создаем новый документ .docx
        XWPFDocument document = new XWPFDocument();

        // Добавляем заголовок в документ
        document.createParagraph().createRun().setText("Отчет" + reportType);

        // Создаем таблицу
        XWPFTable table = document.createTable();

        // Создаем первую строку таблицы с заголовками
        XWPFTableRow tableRow = table.getRow(0);
        tableRow.getCell(0).setText("Header 1");
        tableRow.addNewTableCell().setText("Header 2");
        tableRow.addNewTableCell().setText("Header 3");

        // Добавляем еще одну строку с данными
        XWPFTableRow row = table.createRow();
        row.getCell(0).setText("Data 1");
        row.getCell(1).setText("Data 2");
        row.getCell(2).setText("Data 3");

        // Записываем документ в выходной поток
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.write(outputStream);

        // Возвращаем байты документа
        return outputStream.toByteArray();
    }
}
