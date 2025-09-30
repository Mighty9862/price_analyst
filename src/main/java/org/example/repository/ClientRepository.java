// repository/ClientRepository.java
package org.example.repository;

import org.example.entity.Client;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface ClientRepository extends JpaRepository<Client, Long> {
    boolean existsByInn(String inn);
    Optional<Client> findByInn(String inn);
}