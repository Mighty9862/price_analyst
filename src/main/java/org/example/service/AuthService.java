// service/AuthService.java
package org.example.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.example.dto.RegistrationRequest;
import org.example.entity.Client;
import org.example.repository.ClientRepository;
import org.springframework.stereotype.Service;

@Slf4j
@Service
@RequiredArgsConstructor
public class AuthService {

    private final ClientRepository clientRepository;

    public boolean validateInn(String inn) {
        if (inn == null) return false;

        // Проверка длины ИНН (10 для ООО, 12 для ИП)
        int length = inn.length();
        if (length != 10 && length != 12) {
            return false;
        }

        // Проверка, что строка состоит только из цифр
        if (!inn.matches("\\d+")) {
            return false;
        }

        return true;
    }

    public Client registerClient(RegistrationRequest request) {
        if (!validateInn(request.getInn())) {
            throw new IllegalArgumentException("Неверный формат ИНН");
        }

        if (clientRepository.existsByInn(request.getInn())) {
            throw new IllegalArgumentException("Клиент с таким ИНН уже зарегистрирован");
        }

        Client client = Client.builder()
                .inn(request.getInn())
                .username(request.getUsername())
                .phone(request.getPhone())
                .build();

        return clientRepository.save(client);
    }
}