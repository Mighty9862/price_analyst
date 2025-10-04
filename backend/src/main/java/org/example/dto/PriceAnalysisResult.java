package org.example.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class PriceAnalysisResult {
    private String barcode;
    private Integer quantity;
    private String bestSupplierName;
    private String bestSupplierSap;
    private Double bestPrice;
    private String productName;
    private Boolean requiresManualProcessing;

    private Integer supplierQuantity; // Новое поле: количество у поставщика
    private Double totalPrice; // Новое поле: общая сумма (quantity * bestPrice)
}