package com.example;

import com.example.service.ExcelService;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class ExcelApplication {
    
    public static void main(String[] args) {
        SpringApplication.run(ExcelApplication.class, args);
    }

    @Bean
    public CommandLineRunner commandLineRunner(ExcelService excelService) {
        return args -> {
            try {
                excelService.updateExcel("C:/tmp/20250210.xlsx");
                System.out.println("Excel file updated successfully");
            } catch (Exception e) {
                System.err.println("Error updating Excel file: " + e.getMessage());
                e.printStackTrace();
            }
        };
    }
} 