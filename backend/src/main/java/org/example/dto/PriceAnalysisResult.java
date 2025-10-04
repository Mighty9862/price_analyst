package org.example.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class PriceAnalysisResult {
    private String barcode;
    private Integer quantity;
    private String productName;
    private Boolean requiresManualProcessing;
    private List<SupplierDetail> bestSuppliers; // Новый список для нескольких поставщиков
    private Double totalPrice; // Общая сумма
    private String message; // Дополнительное сообщение, если не хватает количества

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class SupplierDetail {
        private String supplierName;
        private Double price;
        private Integer quantityTaken; // Сколько берем от этого поставщика
        private Integer supplierQuantity; // Сколько всего у поставщика
    }
}