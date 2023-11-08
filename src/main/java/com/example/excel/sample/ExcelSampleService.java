package com.example.excel.sample;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
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
        cell0.setCellValue("과목");

        Cell cell1 = headerRow.createCell(1);
        cell1.setCellValue("학년");

        Row dataRow = sheet.createRow(1);
        Cell dataCell0 = dataRow.createCell(0);
        dataCell0.setCellValue("수학");

        Cell dataCell1 = dataRow.createCell(1);
        dataCell1.setCellValue("1학년");

        return workbook;
    }

    public Workbook createExcelWithValidation() {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sample Sheet");

        // 엑셀 파일 생성 및 데이터 셋팅 로직
        Row headerRow = sheet.createRow(0);
        Cell cell0 = headerRow.createCell(0);
        cell0.setCellValue("과목");

        Row dataRow = sheet.createRow(1);
        Cell dataCell0 = dataRow.createCell(0);
        dataCell0.setCellValue(""); // 빈 셀로 초기화

        int lastRowIndex = sheet.getLastRowNum();

        // 데이터 유효성 검사 설정
        DataValidationHelper validationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint validationConstraint = validationHelper.createExplicitListConstraint(new String[]{"수학", "과학"});
        CellRangeAddressList addressList = new CellRangeAddressList(1, lastRowIndex, 0, 0);
        DataValidation validation = validationHelper.createValidation(validationConstraint, addressList);
        sheet.addValidationData(validation);

        return workbook;
    }


    public Workbook createExcelWithValidationAndVbaExcelFile() {
        try {
            // 기존 엑셀 파일 읽기
            FileInputStream fileInputStream = new FileInputStream("src/main/resources/excelFiles/샘플파일.xlsm");
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            String[] validationArr = new String[]{"수학", "과학", "영어"};

            // 데이터 유효성 검사 설정
            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            DataValidationConstraint validationConstraint = validationHelper.createExplicitListConstraint(validationArr);
            CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
            DataValidation validation = validationHelper.createValidation(validationConstraint, addressList);
            sheet.addValidationData(validation);

            // 변경된 엑셀 파일 저장
            FileOutputStream fileOutputStream = new FileOutputStream("새로운_엑셀_파일.xlsm");
            workbook.write(fileOutputStream);

            return workbook;
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

}
