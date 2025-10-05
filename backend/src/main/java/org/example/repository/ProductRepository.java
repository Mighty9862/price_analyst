package org.example.repository;

import org.example.entity.Product;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface ProductRepository extends JpaRepository<Product, Long> {

    // Базовый поиск по списку штрихкодов
    List<Product> findByBarcodeIn(List<String> barcodes);

    // Оптимизированный поиск лучших цен для одного штрихкода
    @Query("SELECT p FROM Product p WHERE p.barcode = :barcode ORDER BY p.priceWithVat ASC")
    List<Product> findBestPricesByBarcode(@Param("barcode") String barcode);

    // Оптимизированный поиск лучших цен для списка штрихкодов
    // Возвращает ВСЕ товары для указанных штрихкодов, отсортированные по штрихкоду и цене
    @Query("SELECT p FROM Product p WHERE p.barcode IN :barcodes ORDER BY p.barcode, p.priceWithVat ASC")
    List<Product> findBestPricesByBarcodes(@Param("barcodes") List<String> barcodes);

    boolean existsByBarcode(String barcode);

    Optional<Product> findBySupplier_SupplierNameAndBarcode(String supplierName, String barcode);

    boolean existsBySupplier_SupplierNameAndBarcode(String supplierName, String barcode);

    // Статистика по количеству товаров
    @Query("SELECT COUNT(p) FROM Product p")
    long countProducts();
}