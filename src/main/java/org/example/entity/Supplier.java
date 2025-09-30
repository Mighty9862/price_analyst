// entity/Supplier.java
package org.example.entity;

import jakarta.persistence.*;
import lombok.*;

@Entity
@Table(name = "suppliers")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Supplier {
    @Id
    @Column(name = "supplier_sap")
    private String supplierSap;

    @Column(nullable = false)
    private String supplierName;
}