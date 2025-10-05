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
            int barcodeCol = findColumnIndex(sheet, "Штрихкод");
            int quantityCol = findColumnIndex(sheet, "Количество");

            // Проверяем, что колонки найдены
            if (barcodeCol == -1 || quantityCol == -1) {
                throw new IllegalArgumentException("Не найдены необходимые заголовки 'Штрихкод' или 'Количество' в файле.");
            }

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

                if (barcode == null || barcode.trim().isEmpty()) {
                    results.add(createErrorResult(null, quantity, i + 1, "Штрихкод отсутствует или пустой"));
                    continue;
                }

                barcode = barcode.trim();

                if (!isValidBarcode(barcode)) {
                    results.add(createErrorResult(barcode, quantity, i + 1, "Недопустимый формат штрихкода: " + barcode));
                    continue;
                }

                if (quantity == null || quantity <= 0) {
                    results.add(createErrorResult(barcode, quantity, i + 1, "Количество должно быть больше нуля"));
                    continue;
                }

                barcodes.add(barcode);
                barcodeQuantities.put(barcode, quantity);
                rowDataList.add(new RowData(barcode, quantity, i));
            }

            log.info("Found {} unique barcodes in file", barcodes.size());

            if (barcodes.isEmpty()) {
                log.warn("No valid barcodes found in the file");
                return results;
            }

            // Оптимизированный поиск всех цен за один запрос
            List<Product> bestProducts = productRepository.findBestPricesByBarcodes(new ArrayList<>(barcodes));

            // Логируем первые 3 найденных товара для отладки
            if (!bestProducts.isEmpty()) {
                log.info("Первые 3 товара из базы данных:");
                for (int i = 0; i < Math.min(3, bestProducts.size()); i++) {
                    Product p = bestProducts.get(i);
                    log.info("Товар {}: Штрихкод={}, Наименование={}, Поставщик={}, Цена={}, Количество={}",
                            i + 1, p.getBarcode(), p.getProductName(),
                            p.getSupplier().getSupplierName(), p.getPriceWithVat(), p.getQuantity());
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
                        continue;
                    }

                    // Сортируем по цене ascending
                    List<Product> sortedProducts = products.stream()
                            .sorted(Comparator.comparing(Product::getPriceWithVat))
                            .toList();

                    // Проверяем, есть ли поставщики с ненулевым количеством
                    boolean hasAvailableProducts = sortedProducts.stream()
                            .anyMatch(p -> p.getQuantity() != null && p.getQuantity() > 0);

                    if (!hasAvailableProducts) {
                        results.add(createManualProcessingResult(rowData.barcode, rowData.quantity,
                                "Нет доступного количества у поставщиков для товара"));
                        continue;
                    }

                    // Жадно набираем количество от самых дешевых
                    List<PriceAnalysisResult.SupplierDetail> supplierDetails = new ArrayList<>();
                    int remainingQuantity = rowData.quantity;
                    double totalPrice = 0.0;
                    String productName = sortedProducts.get(0).getProductName(); // Берем имя от первого (они должны быть одинаковыми)
                    StringBuilder messageBuilder = new StringBuilder();
                    boolean enoughQuantity = true;

                    for (Product p : sortedProducts) {
                        if (remainingQuantity <= 0) break;

                        if (p.getQuantity() == null || p.getQuantity() <= 0) continue;

                        int take = Math.min(remainingQuantity, p.getQuantity());
                        supplierDetails.add(PriceAnalysisResult.SupplierDetail.builder()
                                .supplierName(p.getSupplier().getSupplierName())
                                .price(p.getPriceWithVat())
                                .quantityTaken(take)
                                .supplierQuantity(p.getQuantity())
                                .build());

                        totalPrice += take * p.getPriceWithVat();
                        remainingQuantity -= take;
                    }

                    if (remainingQuantity > 0) {
                        // Не набрали, но выводим что есть
                        enoughQuantity = false;
                    }

                    // Формируем сообщение в зависимости от ситуации
                    if (enoughQuantity && supplierDetails.size() == 1) {
                        // Успешно взяли всё у одного поставщика
                        messageBuilder.append(String.format("Взято %d шт. у поставщика %s по цене %.2f за единицу",
                                rowData.quantity, supplierDetails.get(0).getSupplierName(), supplierDetails.get(0).getPrice()));
                    } else if (enoughQuantity && supplierDetails.size() > 1) {
                        // Взяли у нескольких поставщиков
                        messageBuilder.append("Взято у нескольких поставщиков: ");
                        for (int i = 0; i < supplierDetails.size(); i++) {
                            PriceAnalysisResult.SupplierDetail detail = supplierDetails.get(i);
                            messageBuilder.append(String.format("%d шт. у %s по цене %.2f",
                                    detail.getQuantityTaken(), detail.getSupplierName(), detail.getPrice()));
                            if (i < supplierDetails.size() - 1) {
                                messageBuilder.append("; ");
                            }
                        }
                    } else if (!enoughQuantity && !supplierDetails.isEmpty()) {
                        // Недостаточно количества, взяли что есть
                        int takenQuantity = rowData.quantity - remainingQuantity;
                        messageBuilder.append(String.format("Недостаточно количества. Доступно только %d из %d шт. Взято: ",
                                takenQuantity, rowData.quantity));
                        for (int i = 0; i < supplierDetails.size(); i++) {
                            PriceAnalysisResult.SupplierDetail detail = supplierDetails.get(i);
                            messageBuilder.append(String.format("%d шт. у %s по цене %.2f",
                                    detail.getQuantityTaken(), detail.getSupplierName(), detail.getPrice()));
                            if (i < supplierDetails.size() - 1) {
                                messageBuilder.append("; ");
                            }
                        }
                        messageBuilder.append(String.format(". Не хватает %d шт.", remainingQuantity));
                    } else {
                        // Нет доступного количества вообще
                        messageBuilder.append("Нет доступного количества у поставщиков для товара");
                    }

                    PriceAnalysisResult result = PriceAnalysisResult.builder()
                            .barcode(rowData.barcode)
                            .quantity(rowData.quantity)
                            .productName(productName != null ? productName : "Не указано")
                            .bestSuppliers(supplierDetails)
                            .totalPrice(totalPrice > 0 ? totalPrice : null)
                            .requiresManualProcessing(!enoughQuantity || supplierDetails.isEmpty())
                            .message(messageBuilder.toString())
                            .build();

                    results.add(result);
                } catch (Exception e) {
                    log.warn("Ошибка обработки строки {}: {}", rowData.rowNumber + 1, e.getMessage());
                    results.add(createErrorResult(rowData.barcode, rowData.quantity, rowData.rowNumber + 1, e.getMessage()));
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
     * Создает результат для ручной обработки
     */
    private PriceAnalysisResult createManualProcessingResult(String barcode, Integer quantity, String reason) {
        return PriceAnalysisResult.builder()
                .barcode(barcode)
                .quantity(quantity)
                .requiresManualProcessing(true)
                .productName(reason)
                .bestSuppliers(Collections.emptyList())
                .totalPrice(null)
                .message(reason)
                .build();
    }

    /**
     * Создает результат с ошибкой
     */
    private PriceAnalysisResult createErrorResult(String barcode, Integer quantity, int rowNumber, String errorMessage) {
        return PriceAnalysisResult.builder()
                .barcode(barcode)
                .quantity(quantity)
                .requiresManualProcessing(true)
                .productName("Ошибка обработки строки " + rowNumber + ": " + errorMessage)
                .bestSuppliers(Collections.emptyList())
                .totalPrice(null)
                .message("Ошибка обработки строки " + rowNumber + ": " + errorMessage)
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

    private int findColumnIndex(Sheet sheet, String expectedHeader) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return -1;

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            String cellValue = getCellStringValue(headerRow.getCell(i));
            if (cellValue != null && cellValue.trim().equalsIgnoreCase(expectedHeader.trim())) {
                return i;
            }
        }
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

    private Integer getCellIntegerValue(Cell cell) {
        if (cell == null) return null;

        return switch (cell.getCellType()) {
            case NUMERIC -> (int) cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Integer.parseInt(cell.getStringCellValue());
                } catch (NumberFormatException e) {
                    yield null;
                }
            }
            default -> null;
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