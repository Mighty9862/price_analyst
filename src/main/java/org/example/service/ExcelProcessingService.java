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

        // Список для примеров дубликатов
        List<String> duplicateExamples = new ArrayList<>();

        ExcelUploadResponse response = ExcelUploadResponse.builder().build();
        List<String> errors = new ArrayList<>();
        int processed = 0;
        int failed = 0;
        int skipped = 0;

        // Кэш поставщиков для оптимизации
        Map<String, Supplier> supplierCache = new HashMap<>();
        // Кэш для проверки дубликатов ТОЛЬКО В ФАЙЛЕ (supplierSap + barcode)
        Map<String, Boolean> fileDuplicateCheckCache = new HashMap<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            // Определяем индексы колонок с более точной логикой
            int supplierSapCol = findColumnIndex(sheet, "сап поставщик", "sap поставщика", "сап");
            int supplierNameCol = findColumnIndex(sheet, "наименование поставщика", "поставщик");
            int barcodeCol = findColumnIndex(sheet, "штрих", "шк", "barcode", "штрихкод", "штрих код");
            int productSapCol = findColumnIndex(sheet, "сап товара", "sap товара", "товар", "код товара");
            int productNameCol = findColumnIndex(sheet, "наименование", "наименование товара", "продукт", "товар");
            int priceCol = findColumnIndex(sheet, "цена", "пц", "price", "стоимость", "пц с ндс опт");

            log.info("Detected columns - SupplierSAP: {}, SupplierName: {}, Barcode: {}, ProductSAP: {}, ProductName: {}, Price: {}",
                    supplierSapCol, supplierNameCol, barcodeCol, productSapCol, productNameCol, priceCol);

            // Логируем заголовки для отладки
            Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                List<String> headers = new ArrayList<>();
                for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                    String header = getCellStringValue(headerRow.getCell(i));
                    headers.add(header != null ? header : "empty");
                }
                log.info("Actual headers in file: {}", headers);
            }

            List<Product> batchProducts = new ArrayList<>();
            int batchSize = 1000;

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                try {
                    String supplierSap = getCellStringValue(row.getCell(supplierSapCol));
                    String barcode = getCellStringValue(row.getCell(barcodeCol));
                    String productName = getCellStringValue(row.getCell(productNameCol));

                    if (supplierSap == null || supplierSap.trim().isEmpty()) {
                        throw new IllegalArgumentException("Не указан SAP код поставщика");
                    }
                    if (barcode == null || barcode.trim().isEmpty()) {
                        throw new IllegalArgumentException("Не указан штрихкод");
                    }

                    supplierSap = supplierSap.trim();
                    barcode = barcode.trim();

                    // Проверка дубликата ТОЛЬКО в текущем файле
                    String duplicateKey = supplierSap + "|" + barcode;
                    if (fileDuplicateCheckCache.containsKey(duplicateKey)) {
                        skipped++;

                        // Сохраняем примеры дубликатов (первые 3)
                        if (duplicateExamples.size() < 3) {
                            String duplicateInfo = String.format("Строка %d: Поставщик %s, Штрихкод %s, Товар %s",
                                    i + 1, supplierSap, barcode, productName != null ? productName : "N/A");
                            duplicateExamples.add(duplicateInfo);
                            log.info("Пример дубликата: {}", duplicateInfo);
                        }

                        log.debug("Пропущен дубликат в файле: поставщик {}, штрихкод {}", supplierSap, barcode);
                        continue;
                    }
                    fileDuplicateCheckCache.put(duplicateKey, true);

                    Product product = processDataRow(row, supplierSapCol, supplierNameCol, barcodeCol,
                            productSapCol, productNameCol, priceCol, supplierCache);
                    if (product != null) {
                        batchProducts.add(product);
                        processed++;

                        // Логируем первые 3 записи для отладки
                        if (processed <= 3) {
                            log.info("Пример сохраненной записи {}: Поставщик={}, Штрихкод={}, Товар={}, Цена={}",
                                    processed, product.getSupplier().getSupplierSap(),
                                    product.getBarcode(), product.getProductName(), product.getPriceWithVat());
                        }

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

            // Формируем сообщение с примерами дубликатов
            String message;
            if (!duplicateExamples.isEmpty()) {
                message = String.format("Обработано записей: %d, пропущено дубликатов В ФАЙЛЕ: %d, ошибок: %d. Время выполнения: %d мс. Примеры дубликатов: %s",
                        processed, skipped, failed, (System.currentTimeMillis() - startTime), String.join("; ", duplicateExamples));
            } else {
                message = String.format("Обработано записей: %d, пропущено дубликатов В ФАЙЛЕ: %d, ошибок: %d. Время выполнения: %d мс",
                        processed, skipped, failed, (System.currentTimeMillis() - startTime));
            }

            response.setSuccess(true);
            response.setMessage(message);
            response.setProcessedRecords(processed);
            response.setFailedRecords(failed);
            response.setDuplicateExamples(duplicateExamples); // Устанавливаем примеры дубликатов

            log.info("File processing completed. Total: {}, Success: {}, Skipped duplicates in file: {}, Failed: {}, Time: {} ms",
                    processed + skipped + failed, processed, skipped, failed, (System.currentTimeMillis() - startTime));

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

        supplierSap = supplierSap.trim();
        barcode = barcode.trim();

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

        // Создаем продукт с правильными данными
        Product product = Product.builder()
                .supplier(supplier)
                .barcode(barcode)
                .productSap(productSap)
                .productName(productName) // Сохраняем правильное имя продукта
                .priceWithVat(price)
                .build();

        // Логируем для отладки
        if (product.getProductName() == null || product.getProductName().equals(supplierName)) {
            log.warn("Возможная проблема с именем продукта: supplierName={}, productName={}",
                    supplierName, productName);
        }

        return product;
    }

    private int findColumnIndex(Sheet sheet, String... keywords) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return 0;

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            String cellValue = getCellStringValue(headerRow.getCell(i));
            if (cellValue != null) {
                String lowerValue = cellValue.toLowerCase();
                for (String keyword : keywords) {
                    if (lowerValue.contains(keyword.toLowerCase())) {
                        log.debug("Найдена колонка '{}' по ключевому слову '{}' в позиции {}", cellValue, keyword, i);
                        return i;
                    }
                }
            }
        }

        // Если не нашли, возвращаем дефолтные значения на основе типичной структуры
        log.warn("Не удалось определить колонки по ключевым словам, используем дефолтные индексы");
        return getDefaultColumnIndex(keywords);
    }

    private int getDefaultColumnIndex(String... keywords) {
        // Дефолтные индексы на основе типичной структуры файла
        String firstKeyword = keywords[0].toLowerCase();
        if (firstKeyword.contains("сап") && firstKeyword.contains("поставщик")) return 0;
        if (firstKeyword.contains("наименование") && firstKeyword.contains("поставщик")) return 1;
        if (firstKeyword.contains("штрих") || firstKeyword.contains("шк")) return 2;
        if (firstKeyword.contains("сап") && !firstKeyword.contains("поставщик")) return 3;
        if (firstKeyword.contains("наименование") && !firstKeyword.contains("поставщик")) return 5; // Исправлено на 5
        if (firstKeyword.contains("цена") || firstKeyword.contains("пц")) return 6; // Исправлено на 6
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