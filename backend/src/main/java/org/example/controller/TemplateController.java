// controller/TemplateController.java
package org.example.controller;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("/api/template")
@RequiredArgsConstructor
@Tag(name = "Шаблоны", description = "API для работы с шаблонами файлов")
public class TemplateController {

    @GetMapping("/download")
    @Operation(summary = "Скачать шаблон", description = "Скачать шаблон Excel файла для загрузки данных")
    public void downloadTemplate(HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=price_analysis_template.xlsx");

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Шаблон для анализа цен");

            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            Row headerRow = sheet.createRow(0);
            String[] headers = {"Штрихкод", "Количество"};

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            Row exampleRow1 = sheet.createRow(1);
            exampleRow1.createCell(0).setCellValue("4606068663735");
            exampleRow1.createCell(1).setCellValue(10);

            Row exampleRow2 = sheet.createRow(2);
            exampleRow2.createCell(0).setCellValue("4600905000264");
            exampleRow2.createCell(1).setCellValue(2);

            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(response.getOutputStream());
        }
    }
}