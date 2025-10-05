package org.example.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelUploadResponse {
    private boolean success;
    private String message;
    private int newRecords;
    private int updatedRecords;
    private int unchangedRecords;
    private int skippedRecords;
    private int processedRecords;
    private int failedRecords;
    private List<String> duplicateExamples; // Примеры дубликатов
}