package com.rosing.demo;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOError;
import java.io.IOException;


public class Test {

    public static void main(String[] args)throws IOException {
        readExcel(2);


    }

    public static void readExcel(int headRowNumber)throws IOException{
        File file = new File("d://tmp/aaaaaa.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheetAt(0);

        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        System.out.println("first row num="+firstRowNum);
        System.out.println("last row num="+lastRowNum);
        Row headRow = sheet.getRow(headRowNumber-1);
        short firstCellNum = headRow.getFirstCellNum();
        short lastCellNum = headRow.getLastCellNum();
        System.out.println("first column num="+firstCellNum);
        System.out.println("last column num="+lastCellNum);

        readRow(headRow);


        for(int i= (headRowNumber);i< lastRowNum;i++){
            Row row = sheet.getRow(i);

        }
    }

    private static void readRow(Row row){
        short firstCellNum = row.getFirstCellNum();
        short lastCellNum = row.getLastCellNum();
        System.out.println("first column num="+firstCellNum);
        System.out.println("last column num="+lastCellNum);
        DataFormatter dataFormatter = new DataFormatter();
        for(int i=firstCellNum;i<lastCellNum;i++){
            Cell cell = row.getCell(i);
            CellType cellType = cell.getCellTypeEnum();
            String stringCellValue = dataFormatter.formatCellValue(cell);
            System.out.println(stringCellValue);
        }

    }

}
