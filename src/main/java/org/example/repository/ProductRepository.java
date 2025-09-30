package org.example.repository;

import org.example.entity.Product;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Set;

@Repository
public interface ProductRepository extends JpaRepository<Product, Long> {

    // Базовый поиск по списку штрихкодов
    List<Product> findByBarcodeIn(List<String> barcodes);

    // Оптимизированный поиск лучших цен для одного штрихкода
    @Query("SELECT p FROM Product p WHERE p.barcode = :barcode ORDER BY p.priceWithVat ASC LIMIT 10")
    List<Product> findBestPricesByBarcode(@Param("barcode") String barcode);

    // Оптимизированный поиск лучших цен для списка штрихкодов
    @Query(value = """
        SELECT DISTINCT ON (p.barcode) p.* 
        FROM products p 
        WHERE p.barcode IN :barcodes 
        ORDER BY p.barcode, p.price_with_vat ASC
        """, nativeQuery = true)
    List<Product> findBestPricesByBarcodes(@Param("barcodes") List<String> barcodes);

    // Пакетный поиск с лимитом
    @Query("SELECT p FROM Product p WHERE p.barcode IN :barcodes ORDER BY p.priceWithVat ASC")
    List<Product> findBestPricesByBarcodesBatch(@Param("barcodes") List<String> barcodes);

    boolean existsByBarcode(String barcode);

    // Статистика по количеству товаров
    @Query("SELECT COUNT(p) FROM Product p")
    long countProducts();

    // Поиск по диапазону ID для пакетной обработки
    @Query("SELECT p FROM Product p WHERE p.id BETWEEN :startId AND :endId")
    List<Product> findProductsByIdRange(@Param("startId") Long startId, @Param("endId") Long endId);
}