package com.backend.excel.service.impl;

import com.backend.excel.model.Product;
import com.backend.excel.repository.ProductRepository;
import com.backend.excel.service.ProductService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ProductServiceImpl implements ProductService {

    private final ProductRepository productRepository;

    // 🔹 Excel import (Excel → DB)
    @Override
    public void importExcel(MultipartFile file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
            List<Product> products = new ArrayList<>();

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Product p = Product.builder()
                        .name(getStringValue(row.getCell(0)))
                        .price(getNumericValue(row.getCell(1)))
                        .quantity((int) getNumericValue(row.getCell(2)))
                        .build();

                products.add(p);
            }

            productRepository.saveAll(products);
            saveExcelToFile(); // importdan sonra avtomatik Excel də yenilənsin
        }
    }

    // 🔹 Excel export (DB → Excel)
    @Override
    public ByteArrayInputStream exportExcel() {
        String[] columns = {"Name", "Price", "Quantity"};

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Products");

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.length; i++) {
                headerRow.createCell(i).setCellValue(columns[i]);
            }

            List<Product> products = productRepository.findAll();
            int rowIdx = 1;
            for (Product p : products) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(p.getName());
                row.createCell(1).setCellValue(p.getPrice());
                row.createCell(2).setCellValue(p.getQuantity());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("Excel export failed: " + e.getMessage());
        }
    }

    // 🔹 Yeni məhsul əlavə (JSON → DB)
    @Override
    public void save(Product product) {
        productRepository.save(product);
        saveExcelToFile(); // hər yeni məhsulda Excel faylı avtomatik yenilənir
    }

    // 🔹 Bütün məhsulları DB-dən götür
    @Override
    public List<Product> getAllProducts() {
        return productRepository.findAll();
    }

    // 🔹 Excel faylını diskə saxlayır
    private void saveExcelToFile() {
        try (FileOutputStream fos = new FileOutputStream("products.xlsx")) {
            fos.write(exportExcel().readAllBytes());
            System.out.println("✅ Excel faylı yeniləndi: products.xlsx");
        } catch (IOException e) {
            System.err.println("⚠ Excel faylı saxlanmadı: " + e.getMessage());
        }
    }

    // 🔹 Boş hüceyrələr üçün təhlükəsiz oxuma metodları
    private String getStringValue(Cell cell) {
        return (cell == null) ? "" : cell.getStringCellValue().trim();
    }

    private double getNumericValue(Cell cell) {
        if (cell == null) return 0;
        if (cell.getCellType() == CellType.NUMERIC) return cell.getNumericCellValue();
        try {
            return Double.parseDouble(cell.getStringCellValue());
        } catch (Exception e) {
            return 0;
        }
    }
}
