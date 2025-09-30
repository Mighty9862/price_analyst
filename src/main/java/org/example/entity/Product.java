package org.example.entity;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.BatchSize;

@Entity
@Table(name = "products", indexes = {
        @Index(name = "idx_barcode", columnList = "barcode"),
        @Index(name = "idx_barcode_price", columnList = "barcode, price_with_vat"),
        @Index(name = "idx_supplier_barcode", columnList = "supplier_sap, barcode")
})
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
@BatchSize(size = 50)
public class Product {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "supplier_sap")
    private Supplier supplier;

    @Column(nullable = false)
    private String barcode;

    private String productSap;
    private String productName;

    @Column(name = "price_with_vat")
    private Double priceWithVat;
}