package org.example.config;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.info.Info;
import io.swagger.v3.oas.models.info.Contact;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class OpenApiConfig {

    @Bean
    public OpenAPI customOpenAPI() {
        return new OpenAPI()
                .info(new Info()
                        .title("Price Analysis Service API")
                        .version("1.0")
                        .description("API для сервиса анализа лучших цен. CORS настроен для всех доменов.")
                        .contact(new Contact()
                                .name("Support Team")
                                .email("support@priceservice.com")));
    }
}