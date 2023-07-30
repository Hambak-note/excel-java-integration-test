package com.example.excel.sample;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ExcelSampleService {

    public Workbook createExcelWorkbook() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sample Sheet");

        // 엑셀 파일 생성 및 데이터 셋팅 로직
        Row headerRow = sheet.createRow(0);
        Cell cell0 = headerRow.createCell(0);
        cell0.setCellValue("이름");

        Cell cell1 = headerRow.createCell(1);
        cell1.setCellValue("나이");

        Row dataRow = sheet.createRow(1);
        Cell dataCell0 = dataRow.createCell(0);
        dataCell0.setCellValue("John");

        Cell dataCell1 = dataRow.createCell(1);
        dataCell1.setCellValue(30);

        return workbook;
    }
}
