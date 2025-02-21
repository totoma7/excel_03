package com.example.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ExcelGeneratorService {
    
    public void generateExcel() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet1 = workbook.createSheet("일일자금수지");
        Sheet sheet2 = workbook.createSheet("자금현황");

        Font font = workbook.createFont();
        font.setFontName("맑은 고딕");

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(font);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setFont(font);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        createDailyCashFlowSheet(sheet1, headerStyle, bodyStyle);
        createFundStatusSheet(sheet2, headerStyle, bodyStyle);

        try (FileOutputStream outputStream = new FileOutputStream("자금_현황_보고서.xlsx")) {
            workbook.write(outputStream);
        }
        workbook.close();
        System.out.println("엑셀 파일 생성 완료: 자금_현황_보고서.xlsx");
    }

    private void createDailyCashFlowSheet(Sheet sheet, CellStyle headerStyle, CellStyle bodyStyle) {
        // 제목 행
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("(2025년 2월 11일 기준)");
        titleCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 7));

        // ... (기존 createDailyCashFlowSheet 메소드의 나머지 내용)
    }

    private void createFundStatusSheet(Sheet sheet, CellStyle headerStyle, CellStyle bodyStyle) {
        // 제목 행
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("2. 자금현황");
        titleCell.setCellStyle(headerStyle);
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 11));

        // ... (기존 createFundStatusSheet 메소드의 나머지 내용)
    }

    private void createMergedCell(Sheet sheet, Row row, int startCol, int endCol, String value, CellStyle style) {
        Cell cell = row.createCell(startCol);
        cell.setCellValue(value);
        cell.setCellStyle(style);
        if (startCol != endCol) {
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(
                row.getRowNum(), row.getRowNum(), startCol, endCol));
        }
    }
} 