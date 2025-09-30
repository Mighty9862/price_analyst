// entity/Client.java
package org.example.entity;

import jakarta.persistence.*;
import lombok.*;
import java.time.LocalDateTime;

@Entity
@Table(name = "clients")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class Client {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(unique = true, nullable = false)
    private String inn;

    @Column(nullable = false)
    private String username;

    @Column(nullable = false)
    private String phone;

    @Builder.Default
    private LocalDateTime registrationDate = LocalDateTime.now();
}