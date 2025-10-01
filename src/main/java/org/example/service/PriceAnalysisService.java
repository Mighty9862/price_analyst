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

            if (barcodes.isEmpty()) {
                log.warn("No valid barcodes found in the file");
                return results;
            }

            // Оптимизированный поиск всех цен за один запрос
            List<Product> bestProducts = productRepository.findBestPricesByBarcodes(new ArrayList<>(barcodes));

            // Логируем первые 5 найденных товара для отладки
            if (!bestProducts.isEmpty()) {
                log.info("Первые 5 товаров из базы данных:");
                for (int i = 0; i < Math.min(5, bestProducts.size()); i++) {
                    Product p = bestProducts.get(i);
                    log.info("Товар {}: Штрихкод={}, Наименование={}, Поставщик={}, Цена={}",
                            i + 1, p.getBarcode(), p.getProductName(),
                            p.getSupplier().getSupplierName(), p.getPriceWithVat());
                }
            }

            // Группируем товары по штрихкодам для обработки нескольких поставщиков
            Map<String, List<Product>> productsByBarcode = bestProducts.stream()
                    .collect(Collectors.groupingBy(Product::getBarcode));

            log.info("Found products for {} barcodes in database", productsByBarcode.size());

            // Формируем результаты
            for (RowData rowData : rowDataList) {
                try {
                    List<Product> products = productsByBarcode.get(rowData.barcode);

                    if (products == null || products.isEmpty()) {
                        // Товар не найден в базе
                        results.add(createManualProcessingResult(rowData.barcode, rowData.quantity, "Товар не найден в базе"));
                    } else if (products.size() == 1) {
                        // Один поставщик для этого штрихкода
                        Product bestProduct = products.get(0);
                        results.add(createSuccessResult(rowData.barcode, rowData.quantity, bestProduct));
                    } else {
                        // Несколько поставщиков для этого штрихкода - выбираем лучшую цену
                        Product bestProduct = findBestPriceProduct(products);
                        results.add(createSuccessResult(rowData.barcode, rowData.quantity, bestProduct));

                        // Логируем информацию о нескольких поставщиках
                        log.debug("Multiple suppliers for barcode {}: {} suppliers, best price: {} from {}",
                                rowData.barcode, products.size(), bestProduct.getPriceWithVat(),
                                bestProduct.getSupplier().getSupplierName());
                    }
                } catch (Exception e) {
                    log.warn("Ошибка обработки строки {}: {}", rowData.rowNumber + 1, e.getMessage());
                    results.add(createErrorResult(rowData.barcode, rowData.quantity, e));
                }
            }

            // Логируем статистику по нескольким поставщикам
            logMultipleSuppliersStatistics(productsByBarcode);

            long endTime = System.currentTimeMillis();
            log.info("Price analysis completed in {} ms for {} items. Processed {} barcodes with multiple suppliers.",
                    (endTime - startTime), results.size(), countBarcodesWithMultipleSuppliers(productsByBarcode));

        } catch (Exception e) {
            log.error("Ошибка обработки файла пользователя", e);
            throw new RuntimeException("Ошибка обработки файла: " + e.getMessage());
        }

        return results;
    }

    /**
     * Находит товар с лучшей ценой из списка товаров от разных поставщиков
     */
    private Product findBestPriceProduct(List<Product> products) {
        return products.stream()
                .min(Comparator.comparing(Product::getPriceWithVat))
                .orElse(products.get(0));
    }

    /**
     * Создает результат для успешного анализа
     */
    private PriceAnalysisResult createSuccessResult(String barcode, Integer quantity, Product product) {
        // Проверяем, что имя продукта не равно имени поставщика
        String actualProductName = product.getProductName();
        String supplierName = product.getSupplier().getSupplierName();

        if (actualProductName == null || actualProductName.equals(supplierName)) {
            log.warn("Проблема с именем продукта для штрихкода {}: productName='{}', supplierName='{}'",
                    barcode, actualProductName, supplierName);
            // Если имя продукта совпадает с именем поставщика, используем запасной вариант
            actualProductName = "Наименование не указано";
        }

        return PriceAnalysisResult.builder()
                .barcode(barcode)
                .quantity(quantity)
                .bestSupplierName(supplierName)
                .bestSupplierSap(product.getSupplier().getSupplierSap())
                .bestPrice(product.getPriceWithVat())
                .productName(actualProductName)
                .requiresManualProcessing(false)
                .build();
    }

    /**
     * Создает результат для ручной обработки
     */
    private PriceAnalysisResult createManualProcessingResult(String barcode, Integer quantity, String reason) {
        return PriceAnalysisResult.builder()
                .barcode(barcode)
                .quantity(quantity)
                .requiresManualProcessing(true)
                .productName(reason)
                .build();
    }

    /**
     * Создает результат с ошибкой
     */
    private PriceAnalysisResult createErrorResult(String barcode, Integer quantity, Exception e) {
        return PriceAnalysisResult.builder()
                .barcode(barcode)
                .quantity(quantity)
                .requiresManualProcessing(true)
                .productName("Ошибка обработки: " + e.getMessage())
                .build();
    }

    /**
     * Логирует статистику по штрихкодам с несколькими поставщиками
     */
    private void logMultipleSuppliersStatistics(Map<String, List<Product>> productsByBarcode) {
        Map<Integer, Integer> supplierCountStats = new HashMap<>();

        for (List<Product> products : productsByBarcode.values()) {
            int supplierCount = products.size();
            supplierCountStats.merge(supplierCount, 1, Integer::sum);
        }

        // Логируем статистику только если есть штрихкоды с несколькими поставщиками
        supplierCountStats.entrySet().stream()
                .filter(entry -> entry.getKey() > 1)
                .forEach(entry ->
                        log.info("Штрихкодов с {} поставщиками: {}", entry.getKey(), entry.getValue()));
    }

    /**
     * Подсчитывает количество штрихкодов с несколькими поставщиками
     */
    private long countBarcodesWithMultipleSuppliers(Map<String, List<Product>> productsByBarcode) {
        return productsByBarcode.values().stream()
                .filter(products -> products.size() > 1)
                .count();
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