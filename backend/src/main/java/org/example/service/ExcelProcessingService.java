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
        // Кэш для проверки дубликатов ТОЛЬКО В ФАЙЛЕ (supplierName + productName)
        Map<String, Boolean> fileDuplicateCheckCache = new HashMap<>();

        try (Workbook workbook = WorkbookFactory.create(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            // Определяем индексы колонок по фиксированным заголовкам
            int supplierNameCol = findColumnIndex(sheet, "Наименование поставщика");
            int barcodeCol = findColumnIndex(sheet, "Штрих код");
            int externalProductCodeCol = findColumnIndex(sheet, "Товар");
            int productNameCol = findColumnIndex(sheet, "Наименование");
            int priceCol = findColumnIndex(sheet, "ПЦ с НДС опт");
            int quantityCol = findColumnIndex(sheet, "Количество");

            // Проверяем, что все колонки найдены
            if (supplierNameCol == -1 || barcodeCol == -1 ||
                    externalProductCodeCol == -1 || productNameCol == -1 ||
                    priceCol == -1 || quantityCol == -1) {
                throw new IllegalArgumentException("Не найдены все необходимые заголовки в файле. Проверьте формат.");
            }

            log.info("Detected columns - SupplierName: {}, Barcode: {}, ExternalCode: {}, ProductName: {}, Price: {}, Quantity: {}",
                    supplierNameCol, barcodeCol, externalProductCodeCol, productNameCol, priceCol, quantityCol);

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
                    String supplierName = getCellStringValue(row.getCell(supplierNameCol));
                    String barcode = getCellStringValue(row.getCell(barcodeCol));
                    String productName = getCellStringValue(row.getCell(productNameCol));

                    if (supplierName == null || supplierName.trim().isEmpty()) {
                        throw new IllegalArgumentException("Не указано наименование поставщика");
                    }
                    if (barcode == null || barcode.trim().isEmpty()) {
                        throw new IllegalArgumentException("Не указан штрихкод");
                    }

                    supplierName = supplierName.trim();
                    barcode = barcode.trim();
                    if (productName != null) {
                        productName = productName.trim();
                    }

                    // Проверка дубликата ТОЛЬКО в текущем файле (supplierName + productName)
                    if (productName == null || productName.isEmpty()) {
                        // Если имя товара пустое, не проверяем на дубликат
                        log.debug("Имя товара пустое для поставщика {}, штрихкода {}, пропускаем проверку дубликатов", supplierName, barcode);
                    } else {
                        String duplicateKey = supplierName + "|" + productName;
                        if (fileDuplicateCheckCache.containsKey(duplicateKey)) {
                            skipped++;

                            // Сохраняем примеры дубликатов (первые 3)
                            if (duplicateExamples.size() < 3) {
                                String duplicateInfo = String.format("Строка %d: Поставщик %s, Товар %s, Штрихкод %s",
                                        i + 1, supplierName, productName, barcode != null ? barcode : "N/A");
                                duplicateExamples.add(duplicateInfo);
                                log.info("Пример дубликата: {}", duplicateInfo);
                            }

                            log.debug("Пропущен дубликат в файле: поставщик {}, товар {}", supplierName, productName);
                            continue;
                        }
                        fileDuplicateCheckCache.put(duplicateKey, true);
                    }

                    Product product = processDataRow(row, supplierNameCol, barcodeCol,
                            externalProductCodeCol, productNameCol, priceCol, quantityCol, supplierCache);
                    if (product != null) {
                        batchProducts.add(product);
                        processed++;

                        // Логируем первые 3 записи для отладки
                        if (processed <= 3) {
                            log.info("Пример сохраненной записи {}: Поставщик={}, Штрихкод={}, Товар={}, Цена={}, Количество={}",
                                    processed, product.getSupplier().getSupplierName(),
                                    product.getBarcode(), product.getProductName(), product.getPriceWithVat(), product.getQuantity());
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

    private Product processDataRow(Row row, int supplierNameCol, int barcodeCol,
                                   int externalProductCodeCol, int productNameCol, int priceCol,
                                   int quantityCol, Map<String, Supplier> supplierCache) {
        String supplierName = getCellStringValue(row.getCell(supplierNameCol));
        String barcode = getCellStringValue(row.getCell(barcodeCol));
        String externalCode = getCellStringValue(row.getCell(externalProductCodeCol));
        String productName = getCellStringValue(row.getCell(productNameCol));
        Double price = getCellNumericValue(row.getCell(priceCol));
        Integer quantity = getCellIntegerValue(row.getCell(quantityCol));

        if (supplierName == null || supplierName.trim().isEmpty()) {
            throw new IllegalArgumentException("Не указано наименование поставщика");
        }
        if (barcode == null || barcode.trim().isEmpty()) {
            throw new IllegalArgumentException("Не указан штрихкод");
        }

        supplierName = supplierName.trim();
        barcode = barcode.trim();

        if (!isValidBarcode(barcode)) {
            throw new IllegalArgumentException("Неверный формат штрихкода: " + barcode);
        }

        // Используем кэш поставщиков для оптимизации
        Supplier supplier = supplierCache.get(supplierName);
        if (supplier == null) {
            supplier = supplierRepository.findById(supplierName)
                    .orElse(Supplier.builder().supplierName(supplierName).build());
            supplierRepository.save(supplier);
            supplierCache.put(supplierName, supplier);
        }

        // Создаем продукт с правильными данными
        Product product = Product.builder()
                .supplier(supplier)
                .barcode(barcode)
                .externalCode(externalCode)
                .productName(productName) // Сохраняем правильное имя продукта
                .priceWithVat(price)
                .quantity(quantity != null ? quantity : 0)
                .build();

        // Логируем для отладки
        if (product.getProductName() == null || product.getProductName().equals(supplierName)) {
            log.warn("Возможная проблема с именем продукта: supplierName={}, productName={}",
                    supplierName, productName);
        }

        return product;
    }

    private int findColumnIndex(Sheet sheet, String expectedHeader) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return -1;

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            String cellValue = getCellStringValue(headerRow.getCell(i));
            if (cellValue != null && cellValue.trim().equalsIgnoreCase(expectedHeader.trim())) {
                log.debug("Найдена колонка '{}' в позиции {}", cellValue, i);
                return i;
            }
        }

        log.warn("Не найдена колонка с заголовком '{}'", expectedHeader);
        return -1;
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
}