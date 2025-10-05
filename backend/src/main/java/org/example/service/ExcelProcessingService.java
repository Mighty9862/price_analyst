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
import java.util.Optional;
import java.util.Objects;

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
        int newRecords = 0;
        int updatedRecords = 0;
        int unchangedRecords = 0;
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

                    // Получаем данные из строки
                    String externalCode = getCellStringValue(row.getCell(externalProductCodeCol));
                    Double price = getCellNumericValue(row.getCell(priceCol));
                    Integer quantity = getCellIntegerValue(row.getCell(quantityCol));

                    // Используем кэш поставщиков для оптимизации
                    Supplier supplier = supplierCache.get(supplierName);
                    if (supplier == null) {
                        supplier = supplierRepository.findById(supplierName)
                                .orElse(Supplier.builder().supplierName(supplierName).build());
                        supplierRepository.save(supplier);
                        supplierCache.put(supplierName, supplier);
                    }

                    // Проверяем существование продукта в БД по supplierName и barcode
                    Optional<Product> existingProduct = productRepository.findBySupplier_SupplierNameAndBarcode(supplierName, barcode);

                    if (existingProduct.isPresent()) {
                        Product existing = existingProduct.get();
                        // Проверяем, изменились ли данные
                        boolean changed = !Objects.equals(existing.getExternalCode(), externalCode) ||
                                !Objects.equals(existing.getProductName(), productName) ||
                                !Objects.equals(existing.getPriceWithVat(), price) ||
                                !Objects.equals(existing.getQuantity(), quantity != null ? quantity : 0);

                        if (changed) {
                            // Обновляем, если изменились
                            existing.setExternalCode(externalCode);
                            existing.setProductName(productName);
                            existing.setPriceWithVat(price);
                            existing.setQuantity(quantity != null ? quantity : 0);
                            batchProducts.add(existing);
                            updatedRecords++;
                            log.debug("Обновлен существующий продукт: поставщик {}, штрихкод {}", supplierName, barcode);
                        } else {
                            // Без изменений
                            unchangedRecords++;
                            log.debug("Пропущен без изменений: поставщик {}, штрихкод {}", supplierName, barcode);
                        }
                    } else {
                        // Если не существует, создаем новый
                        Product newProduct = Product.builder()
                                .supplier(supplier)
                                .barcode(barcode)
                                .externalCode(externalCode)
                                .productName(productName)
                                .priceWithVat(price)
                                .quantity(quantity != null ? quantity : 0)
                                .build();
                        batchProducts.add(newProduct);
                        newRecords++;
                        log.debug("Добавлен новый продукт: поставщик {}, штрихкод {}", supplierName, barcode);
                    }

                    // Логируем первые 3 записи для отладки (только новые и обновленные)
                    int totalProcessed = newRecords + updatedRecords;
                    if (totalProcessed <= 3) {
                        log.info("Пример обработанной записи {}: Поставщик={}, Штрихкод={}, Товар={}, Цена={}, Количество={}",
                                totalProcessed, supplierName, barcode, productName, price, quantity);
                    }

                    // Пакетное сохранение
                    if (batchProducts.size() >= batchSize) {
                        productRepository.saveAll(batchProducts);
                        batchProducts.clear();
                        log.info("Processed {} records (new: {}, updated: {})...", newRecords + updatedRecords, newRecords, updatedRecords);
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
            int totalProcessed = newRecords + updatedRecords;
            String message = String.format("Добавлено новых записей: %d, обновлено: %d, без изменений: %d, пропущено дубликатов в файле: %d, ошибок: %d. Время выполнения: %d мс",
                    newRecords, updatedRecords, unchangedRecords, skipped, failed, (System.currentTimeMillis() - startTime));

            response.setSuccess(true);
            response.setMessage(message);
            response.setProcessedRecords(totalProcessed);
            response.setFailedRecords(failed);
            response.setDuplicateExamples(duplicateExamples); // Устанавливаем примеры дубликатов

            log.info("File processing completed. New: {}, Updated: {}, Unchanged: {}, Skipped duplicates in file: {}, Failed: {}, Time: {} ms",
                    newRecords, updatedRecords, unchangedRecords, skipped, failed, (System.currentTimeMillis() - startTime));

        } catch (Exception e) {
            log.error("Ошибка обработки файла", e);
            response.setSuccess(false);
            response.setMessage("Ошибка обработки файла: " + e.getMessage());
        }

        return response;
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