package com.example.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

@Service
public class ExcelService {
    
    public void updateExcel(String filePath) throws Exception {
        try (FileInputStream inputStream = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(inputStream)) {
             
            Sheet sheet = workbook.getSheetAt(0);
            
            // Get the last row number
            int lastRowNum = sheet.getLastRowNum();
            
            // Shift rows down from index 37
            if (lastRowNum >= 37) {
                sheet.shiftRows(37, lastRowNum, 1, true, false);
            }

            // Create new row at index 37
            Row newRow = sheet.createRow(37);
            
            // Set values in cells
            Cell cell01 = newRow.createCell(1); 
            Cell cell02 = newRow.createCell(2); 
            Cell cell03 = newRow.createCell(3); 
            Cell cell04 = newRow.createCell(4); 
            Cell cell = newRow.createCell(5); 
            Cell cell1 = newRow.createCell(6); 
            Cell cell2 = newRow.createCell(7); 
            Cell cell3 = newRow.createCell(8); 
            
            cell.setCellValue(900);
            cell1.setCellValue(7012121210L);
            cell2.setCellValue(400);
            
            // Set formula in cell3 (I37 = G37 - H37)
            cell3.setCellFormula("+H38-G38");  // 38은 Excel에서의 실제 행 번호 (1-based)

            // 기존 병합된 셀 찾기 및 제거 (필요한 경우)
            int numMergedRegions = sheet.getNumMergedRegions();
            for (int i = numMergedRegions - 1; i >= 0; i--) {
                CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
                // B33:B37 영역을 찾아서 제거
                if (mergedRegion.getFirstRow() == 32 && mergedRegion.getLastRow() == 36 && 
                    mergedRegion.getFirstColumn() == 1 && mergedRegion.getLastColumn() == 1) {
                    sheet.removeMergedRegion(i);
                    break;
                }
            }
            
            // 새로운 병합 영역 설정 (B33:B38)
            sheet.addMergedRegion(new CellRangeAddress(32, 37, 1, 1));
            
            // 병합된 셀의 첫 번째 셀에 값 설정
            Row row33 = sheet.getRow(32);
            if (row33 != null) {
                Cell cellB33 = row33.getCell(1);
                if (cellB33 != null) {
                    cellB33.setCellValue("보통예금(외화)");
                }
            }
            
            // Create base cell style with black borders and number format
            CellStyle style = workbook.createCellStyle();
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            
            // Set border colors to black
            style.setTopBorderColor(IndexedColors.BLACK.getIndex());
            style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
            style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
            style.setRightBorderColor(IndexedColors.BLACK.getIndex());
            
            // Create base font (맑은 고딕, 10pt)
            Font baseFont = workbook.createFont();
            baseFont.setFontName("맑은 고딕");
            baseFont.setFontHeightInPoints((short) 10);
            style.setFont(baseFont);
            
            // Set number format with commas for all cells
            DataFormat format = workbook.createDataFormat();
            style.setDataFormat(format.getFormat("###,###,###,###,##0"));
            
            // Create special style for cell3 with negative number format
            CellStyle cell3Style = workbook.createCellStyle();
            cell3Style.cloneStyleFrom(style);  // Copy border styles and base font
            
            // Create font for negative numbers (red color, 맑은 고딕, 10pt)
            Font redFont = workbook.createFont();
            redFont.setFontName("맑은 고딕");
            redFont.setFontHeightInPoints((short) 10);
            redFont.setColor(IndexedColors.RED.getIndex());
            cell3Style.setFont(redFont);
            
            // Set number format for negative numbers with parentheses and commas
            cell3Style.setDataFormat(format.getFormat("###,###,###,###,##0;[Red](###,###,###,###,##0)"));
            
            // Apply styles to cells
            cell.setCellStyle(style);
            cell1.setCellStyle(style);
            cell2.setCellStyle(style);
            cell3.setCellStyle(cell3Style);
            
            // Save the updated workbook
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        }
    }
} 