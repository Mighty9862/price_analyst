package org.example.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.example.dto.ExcelUploadResponse;
import org.example.entity.Product;
import org.example.entity.Supplier;
import org.example.repository.ProductRepository;
import org.example.repository.SupplierRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExcelProcessingService {

    private final ProductRepository productRepository;
    private final SupplierRepository supplierRepository;

    @Transactional
    public ExcelUploadResponse processSupplierDataFile(MultipartFile file) {
        long startTime = System.currentTimeMillis();
        ExcelUploadResponse response = ExcelUploadResponse.builder().build();
        List<String> errors = new ArrayList<>();
        int processed = 0;
        int failed = 0;

        // Кэш поставщиков для оптимизации
        Map<String, Supplier> supplierCache = new HashMap<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            // Определяем индексы колонок
            int supplierSapCol = findColumnIndex(sheet, "сап поставщик", "sap поставщика");
            int supplierNameCol = findColumnIndex(sheet, "наименование поставщика");
            int barcodeCol = findColumnIndex(sheet, "штрих", "шк", "barcode");
            int productSapCol = findColumnIndex(sheet, "сап", "sap товара");
            int productNameCol = findColumnIndex(sheet, "наименование", "товар");
            int priceCol = findColumnIndex(sheet, "цена", "пц", "price");

            log.info("Detected columns - SupplierSAP: {}, SupplierName: {}, Barcode: {}, ProductSAP: {}, ProductName: {}, Price: {}",
                    supplierSapCol, supplierNameCol, barcodeCol, productSapCol, productNameCol, priceCol);

            List<Product> batchProducts = new ArrayList<>();
            int batchSize = 1000;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                try {
                    Product product = processDataRow(row, supplierSapCol, supplierNameCol, barcodeCol,
                            productSapCol, productNameCol, priceCol, supplierCache);
                    if (product != null) {
                        batchProducts.add(product);
                        processed++;

                        // Пакетное сохранение
                        if (batchProducts.size() >= batchSize) {
                            productRepository.saveAll(batchProducts);
                            batchProducts.clear();
                            log.info("Processed {} records...", processed);
                        }
                    }
                } catch (Exception e) {
                    failed++;
                    errors.add("Строка " + (i + 1) + ": " + e.getMessage());
                    if (failed <= 10) { // Логируем только первые 10 ошибок
                        log.warn("Ошибка обработки строки {}: {}", i + 1, e.getMessage());
                    }
                }
            }

            // Сохраняем оставшиеся записи
            if (!batchProducts.isEmpty()) {
                productRepository.saveAll(batchProducts);
            }

            response.setSuccess(true);
            response.setMessage(String.format("Обработано записей: %d, ошибок: %d. Время выполнения: %d мс",
                    processed, failed, (System.currentTimeMillis() - startTime)));
            response.setProcessedRecords(processed);
            response.setFailedRecords(failed);

            log.info("File processing completed. Total: {}, Success: {}, Failed: {}, Time: {} ms",
                    processed + failed, processed, failed, (System.currentTimeMillis() - startTime));

        } catch (Exception e) {
            log.error("Ошибка обработки файла", e);
            response.setSuccess(false);
            response.setMessage("Ошибка обработки файла: " + e.getMessage());
        }

        return response;
    }

    private Product processDataRow(Row row, int supplierSapCol, int supplierNameCol, int barcodeCol,
                                   int productSapCol, int productNameCol, int priceCol,
                                   Map<String, Supplier> supplierCache) {
        String supplierSap = getCellStringValue(row.getCell(supplierSapCol));
        String supplierName = getCellStringValue(row.getCell(supplierNameCol));
        String barcode = getCellStringValue(row.getCell(barcodeCol));
        String productSap = getCellStringValue(row.getCell(productSapCol));
        String productName = getCellStringValue(row.getCell(productNameCol));
        Double price = getCellNumericValue(row.getCell(priceCol));

        if (supplierSap == null || supplierSap.trim().isEmpty()) {
            throw new IllegalArgumentException("Не указан SAP код поставщика");
        }
        if (barcode == null || barcode.trim().isEmpty()) {
            throw new IllegalArgumentException("Не указан штрихкод");
        }

        if (!isValidBarcode(barcode)) {
            throw new IllegalArgumentException("Неверный формат штрихкода: " + barcode);
        }

        // Используем кэш поставщиков для оптимизации
        Supplier supplier = supplierCache.get(supplierSap);
        if (supplier == null) {
            supplier = supplierRepository.findById(supplierSap)
                    .orElse(Supplier.builder().supplierSap(supplierSap).build());
            supplier.setSupplierName(supplierName);
            supplierRepository.save(supplier);
            supplierCache.put(supplierSap, supplier);
        }

        return Product.builder()
                .supplier(supplier)
                .barcode(barcode)
                .productSap(productSap)
                .productName(productName)
                .priceWithVat(price)
                .build();
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

    private Double getCellNumericValue(Cell cell) {
        if (cell == null) return 0.0;

        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Double.parseDouble(cell.getStringCellValue().replace(",", "."));
                } catch (NumberFormatException e) {
                    yield 0.0;
                }
            }
            default -> 0.0;
        };
    }

    private boolean isValidBarcode(String barcode) {
        return barcode != null && barcode.matches("\\d{8,14}");
    }
}