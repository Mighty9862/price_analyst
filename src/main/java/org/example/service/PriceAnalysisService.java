package org.example.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.example.dto.PriceAnalysisResult;
import org.example.entity.Product;
import org.example.repository.ProductRepository;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class PriceAnalysisService {

    private final ProductRepository productRepository;

    public List<PriceAnalysisResult> analyzePrices(MultipartFile file) {
        long startTime = System.currentTimeMillis();
        List<PriceAnalysisResult> results = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            // Определяем индексы колонок
            int barcodeCol = findColumnIndex(sheet, "штрих", "barcode", "шк");
            int quantityCol = findColumnIndex(sheet, "количе", "quantity", "кол-во");

            log.info("Using columns - Barcode: {}, Quantity: {}", barcodeCol, quantityCol);

            // Собираем все штрихкоды из файла
            Set<String> barcodes = new HashSet<>();
            Map<String, Integer> barcodeQuantities = new HashMap<>();
            List<RowData> rowDataList = new ArrayList<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                String barcode = getCellStringValue(row.getCell(barcodeCol));
                Integer quantity = getCellIntegerValue(row.getCell(quantityCol));

                if (barcode != null && !barcode.trim().isEmpty() && isValidBarcode(barcode)) {
                    barcode = barcode.trim();
                    barcodes.add(barcode);
                    barcodeQuantities.put(barcode, quantity);
                    rowDataList.add(new RowData(barcode, quantity, i));
                }
            }

            log.info("Found {} unique barcodes in file", barcodes.size());

            // Оптимизированный поиск всех цен за один запрос
            List<Product> bestProducts = productRepository.findBestPricesByBarcodes(new ArrayList<>(barcodes));
            Map<String, Product> bestProductMap = bestProducts.stream()
                    .collect(Collectors.toMap(Product::getBarcode, p -> p, (p1, p2) -> p1));

            log.info("Found best prices for {} barcodes in database", bestProductMap.size());

            // Формируем результаты
            for (RowData rowData : rowDataList) {
                try {
                    Product bestProduct = bestProductMap.get(rowData.barcode);

                    if (bestProduct == null) {
                        results.add(PriceAnalysisResult.builder()
                                .barcode(rowData.barcode)
                                .quantity(rowData.quantity)
                                .requiresManualProcessing(true)
                                .productName("Товар не найден в базе")
                                .build());
                    } else {
                        results.add(PriceAnalysisResult.builder()
                                .barcode(rowData.barcode)
                                .quantity(rowData.quantity)
                                .bestSupplierName(bestProduct.getSupplier().getSupplierName())
                                .bestSupplierSap(bestProduct.getSupplier().getSupplierSap())
                                .bestPrice(bestProduct.getPriceWithVat())
                                .productName(bestProduct.getProductName())
                                .requiresManualProcessing(false)
                                .build());
                    }
                } catch (Exception e) {
                    log.warn("Ошибка обработки строки {}: {}", rowData.rowNumber + 1, e.getMessage());
                    results.add(PriceAnalysisResult.builder()
                            .barcode(rowData.barcode)
                            .quantity(rowData.quantity)
                            .requiresManualProcessing(true)
                            .productName("Ошибка обработки: " + e.getMessage())
                            .build());
                }
            }

            long endTime = System.currentTimeMillis();
            log.info("Price analysis completed in {} ms for {} items", (endTime - startTime), results.size());

        } catch (Exception e) {
            log.error("Ошибка обработки файла пользователя", e);
            throw new RuntimeException("Ошибка обработки файла: " + e.getMessage());
        }

        return results;
    }

    private int findColumnIndex(Sheet sheet, String... keywords) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return 0;

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            String cellValue = getCellStringValue(headerRow.getCell(i));
            if (cellValue != null) {
                String lowerValue = cellValue.toLowerCase();
                for (String keyword : keywords) {
                    if (lowerValue.contains(keyword)) {
                        return i;
                    }
                }
            }
        }
        return 0;
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return null;

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            case BLANK -> null;
            default -> cell.toString().trim();
        };
    }

    private Integer getCellIntegerValue(Cell cell) {
        if (cell == null) return 0;

        return switch (cell.getCellType()) {
            case NUMERIC -> (int) cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Integer.parseInt(cell.getStringCellValue());
                } catch (NumberFormatException e) {
                    yield 0;
                }
            }
            default -> 0;
        };
    }

    private boolean isValidBarcode(String barcode) {
        return barcode != null && barcode.matches("\\d{8,14}");
    }

    // Вспомогательный класс для хранения данных строки
    private static class RowData {
        String barcode;
        Integer quantity;
        int rowNumber;

        RowData(String barcode, Integer quantity, int rowNumber) {
            this.barcode = barcode;
            this.quantity = quantity;
            this.rowNumber = rowNumber;
        }
    }
}