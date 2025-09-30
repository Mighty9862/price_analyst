// dto/UserUploadRequest.java
package org.example.dto;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

@Data
public class UserUploadRequest {
    private MultipartFile file;
}