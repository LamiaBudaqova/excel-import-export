package com.backend.excel.service;

import com.backend.excel.model.Product;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

public interface ProductService {
    void importExcel(MultipartFile file) throws IOException;
    ByteArrayInputStream exportExcel();
    List<Product> getAllProducts();
    void save(Product product);
}
