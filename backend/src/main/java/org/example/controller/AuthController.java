// controller/AuthController.java
package org.example.controller;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.example.dto.RegistrationRequest;
import org.example.entity.Client;
import org.example.service.AuthService;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@RestController
@RequestMapping("/api/auth")
@RequiredArgsConstructor
@Tag(name = "Аутентификация", description = "API для регистрации и проверки ИНН")
public class AuthController {

    private final AuthService authService;

    @PostMapping("/register")
    @Operation(summary = "Регистрация клиента", description = "Регистрация нового клиента в системе")
    public ResponseEntity<Client> register(@Valid @RequestBody RegistrationRequest request) {
        Client client = authService.registerClient(request);
        return ResponseEntity.ok(client);
    }

    @PostMapping("/validate-inn")
    @Operation(summary = "Проверка ИНН", description = "Проверка валидности ИНН")
    public ResponseEntity<Map<String, Boolean>> validateInn(@RequestBody Map<String, String> request) {
        String inn = request.get("inn");
        boolean isValid = authService.validateInn(inn);
        return ResponseEntity.ok(Map.of("valid", isValid));
    }
}