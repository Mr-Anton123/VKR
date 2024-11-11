package com.example.reportgenerator;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

@Controller
public class ReportController {

    // Генерация и скачивание отчета в формате .docx
    @PostMapping("/generateWordReport")
    @ResponseBody
    public void generateWordReport(
            @RequestParam(value = "parameter1", required = false) String parameter1,
            @RequestParam(value = "parameter2", required = false) String parameter2,
            @RequestParam(value = "parameter3", required = false) String parameter3,
            @RequestParam(value = "parameter4", required = false) String parameter4,
            @RequestParam(value = "parameter5", required = false) String parameter5,
            HttpServletResponse response) throws IOException {

        // Если параметров нет, подставляем дефолтные значения
        if (parameter1 == null || parameter1.isEmpty()) parameter1 = "Нет данных";
        if (parameter2 == null || parameter2.isEmpty()) parameter2 = "Нет данных";
        if (parameter3 == null || parameter3.isEmpty()) parameter3 = "Нет данных";
        if (parameter4 == null || parameter4.isEmpty()) parameter4 = "Нет данных";
        if (parameter5 == null || parameter5.isEmpty()) parameter5 = "Нет данных";

        // Создаем новый Word документ
        try (XWPFDocument document = new XWPFDocument()) {

            // Добавляем заголовок
            XWPFParagraph paragraph = document.createParagraph();
            XWPFRun run = paragraph.createRun();
            run.setText("Отчет по шаблону");

            // Добавляем таблицу с данными
            XWPFTable table = document.createTable();

            // Создание заголовка таблицы
            XWPFTableRow headerRow = table.getRow(0); // Создание первой строки таблицы
            setCellText(headerRow.getCell(0), "Параметр", 12);
            setCellText(headerRow.addNewTableCell(), "Описание", 12);
            setCellText(headerRow.addNewTableCell(), "Значение", 12);

            // Устанавливаем ширину столбцов
            setColumnWidth(table, 0, 2000); // Параметр
            setColumnWidth(table, 1, 3000); // Описание
            setColumnWidth(table, 2, 4000); // Значение

            // Заполнение таблицы данными
            addRow(table, "Номер", "1", parameter1);
            addRow(table, "Название", "2", parameter2);
            addRow(table, "Дата", "3", parameter3);
            addRow(table, "Текст", "4", parameter4);
            addRow(table, "Текст 2", "5", parameter5);

            // Устанавливаем выравнивание таблицы
            table.setTableAlignment(TableRowAlign.CENTER);

            // Устанавливаем заголовок для скачивания .docx файла
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setHeader("Content-Disposition", "attachment; filename=report.docx");

            // Отправляем документ в ответ
            try (OutputStream out = response.getOutputStream()) {
                document.write(out);
            }
        } catch (IOException e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("Ошибка при генерации отчета: " + e.getMessage());
        }
    }

    private void addRow(XWPFTable table, String parameter, String description, String value) {
        XWPFTableRow row = table.createRow();
        setCellText(row.getCell(0), parameter, 12);
        setCellText(row.getCell(1), description, 12);
        setCellText(row.getCell(2), value, 12);
    }

    private void setCellText(XWPFTableCell cell, String text, int fontSize) {
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontSize(fontSize);
    }

    private void setColumnWidth(XWPFTable table, int columnIndex, int width) {
        XWPFTableCell cell = table.getRow(0).getCell(columnIndex);
        if (cell != null) {
            // Настроим ширину столбца
            cell.getCTTc().addNewTcPr().addNewTcW().setW(width);
        }
    }
}
