package com.rosing.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
public class Main {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        List<List<String>> datas = getDatas();
        List<CellRangeAddress> addresses = new ArrayList<>();
        for(int i=0;i< datas.size();i++){
            List<String> itemList = datas.get(i);
            Row row = sheet.createRow(i);
            for(int j=0;j<itemList.size();j++){
                Cell cell = row.createCell(j);
                String value = itemList.get(j);
                String preRowValue = null;
                String preColValue = null;
                if(i > 0){
                    preRowValue = datas.get(i-1).get(j);
                }
                if(j>0){
                    preColValue = datas.get(i).get(j-1);
                }
                if(preRowValue!= null && !preRowValue.equals("") && value.equals(preRowValue)){
                    int rowIndex = i-1;
                    int lastRowIndex = i;
                    int colIndex = j;
                    int lastColIndex = j;
                    boolean shouldCreateNew = true;
                    for(CellRangeAddress cellRangeAddress : addresses){
                        int lastRow = cellRangeAddress.getLastRow();
                        int lastColumn = cellRangeAddress.getLastColumn();
                        int firstColumn = cellRangeAddress.getFirstColumn();
                        if(lastRow == rowIndex && lastColumn == colIndex){
                            cellRangeAddress.setLastRow(lastRowIndex);
                            shouldCreateNew = false;
                            break;
                        }
                    }
                    if(shouldCreateNew){
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex,lastRowIndex,colIndex,lastColIndex);
                        addresses.add(cellRangeAddress);
                    }
                }
                cell.setCellValue(value);
            }
        }
        addresses.stream().forEach(cellRangeAddress -> {
            System.out.println("firstRow:"+cellRangeAddress.getFirstRow()+"-lastRow:"+cellRangeAddress.getLastRow()+"-firstCol:"+ cellRangeAddress.getFirstColumn()+"-lastCol:"+ cellRangeAddress.getLastColumn());
        });
        System.out.println("------------");
        for(int i=0;i< datas.size();i++){
            List<String> itemList = datas.get(i);
            for(int j=0;j< itemList.size();j++){
                String value = itemList.get(j);
                String preValue = null;
                if(j>0){
                    preValue = itemList.get(j-1);
                }
                if(preValue != null && !preValue.equals("")){
                    int preColIndex = j-1;
                    int colIndex = j;
                    int rowIndex = i;
                    boolean shouldCreate = false;
                    if(preValue.equals(value)){
                        shouldCreate = true;
                        for(CellRangeAddress address : addresses){
                            int firstRow = address.getFirstRow();
                            int lastRow = address.getLastRow();
                            int firstColumn = address.getFirstColumn();
                            int lastColumn = address.getLastColumn();
                            if((i <=lastRow && i>=firstRow) && (j<=lastColumn )){
                                shouldCreate = false;
                                continue;
                            }
                            if(rowIndex == firstRow && rowIndex== lastRow){
                                address.setLastColumn(colIndex);
                                shouldCreate = false;
                                break;
                            }
                        }
                    }
                    if(shouldCreate){
                        CellRangeAddress address = new CellRangeAddress(rowIndex,rowIndex,preColIndex,colIndex);
                        addresses.add(address);
                    }
                }
            }
        }

        addresses.stream().forEach(cellRangeAddress -> {
            System.out.println("firstRow:"+cellRangeAddress.getFirstRow()+"-lastRow:"+cellRangeAddress.getLastRow()+"-firstCol:"+ cellRangeAddress.getFirstColumn()+"-lastCol:"+ cellRangeAddress.getLastColumn());
        });
        addresses.stream().forEach(cellRangeAddress -> {
            sheet.addMergedRegion(cellRangeAddress);
        });
        FileOutputStream fileOutputStream = new FileOutputStream(new File("d:/tmp/test.xlsx"));
        workbook.write(fileOutputStream);
    }
    private static List<List<String>> getDatas(){
        List<List<String>> resultList = new ArrayList<>();
        resultList.add(getMap("lisi","zhangsan","zhangsan"));
        resultList.add(getMap("lisi","12","89-03-02"));
        resultList.add(getMap("wangwu","11","89-02-02"));
        resultList.add(getMap("zhangsan","zhangsan","89-02-02"));
        resultList.add(getMap("zhangsan","zhangsan","89-03-02"));
        resultList.add(getMap("wangwu","11","89-03-02"));
        resultList.add(getMap("wangwu","11","89-03-02"));

        return resultList;
    }
    private static List<String> getMap(String name,String age,String birthday){
        List<String> resultList = new ArrayList<>();
        resultList.add(name);
        resultList.add(age);
        resultList.add(birthday);
        return resultList;
    }
}