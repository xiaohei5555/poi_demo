package com.rosing.excel;

import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
import com.rosing.demo.ExcelProperty;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author luoxin
 */
public class ExcelWriter<T> {

    private Class<T> clazz;

    private static FormulaEvaluator evaluator;

    public ExcelWriter(Class<T> clazz){
        this.clazz = clazz;
    }


    public void doWrite(List<T> dataList, List<List<String>> headList, OutputStream outputStream)throws Exception{

        Workbook workbook = new XSSFWorkbook();
        evaluator=workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.createSheet();
        //写入头文件
        for(int i=0;i<headList.size();i++){
            List<String> heads = headList.get(i);
            Row row = sheet.createRow(i);
            for(int j=0;j<heads.size();j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(heads.get(j));
            }
        }

        List<String> cacheList = new ArrayList<>();
        List<MergedData> mergedDataList = new ArrayList<>();
        for(int i=0;i<dataList.size();i++){
            T data = dataList.get(i);
            List<String> headDataList = headList.get(headList.size()-1);
            int rowIndex = headList.size()+i;
            Row row = sheet.createRow(rowIndex);
            List<String> rowDataList = new ArrayList<>();
            for(int j=0;j<headDataList.size();j++){
                int colIndex = j;
                String headData = headDataList.get(colIndex);
                String value = getFieldValue(headData,data);
                Cell cell = row.createCell(colIndex);
                cell.setCellValue(value);
                if(!cacheList.isEmpty()){
                    String preData = cacheList.get(colIndex);
                    if(StrUtil.isNotBlank(preData) && StrUtil.isNotBlank(value) && value.equals(preData)){
                        MergedData mergedData = new MergedData(rowIndex-1,colIndex,rowIndex,colIndex,null);
                        addMergedData(mergedDataList,mergedData);
                    }
                }
                rowDataList.add(value);
            }
            cacheList.clear();
            cacheList.addAll(rowDataList);
        }
        mergedDataList.stream().map(mergedData -> new CellRangeAddress(mergedData.getFirstRowIndex(), mergedData.getLastRowIndex(), mergedData.getFirstColumnIndex(), mergedData.getLastColumnIndex())).forEach(cellRangeAddress -> sheet.addMergedRegion(cellRangeAddress));
        workbook.write(outputStream);
    }

    private void addMergedData(List<MergedData> mergedDataList,MergedData mergedData){
        int newFirstRowIndex = mergedData.getFirstRowIndex();
        int newLastRowIndex = mergedData.getLastRowIndex();
        int newFirstColumnIndex = mergedData.getFirstColumnIndex();
        int newLastColumnIndex = mergedData.getLastColumnIndex();

        boolean shouldAdd = true;
        for(MergedData temp: mergedDataList){
            int firstRowIndex = temp.getFirstRowIndex();
            int lastRowIndex = temp.getLastRowIndex();
            int firstColumnIndex = temp.getFirstColumnIndex();
            if(newFirstColumnIndex == firstColumnIndex && lastRowIndex == newFirstRowIndex){
                temp.setLastRowIndex(newLastRowIndex);
                shouldAdd = false;
                break;
            }
        }
        if (shouldAdd) {
            mergedDataList.add(mergedData);
        }

    }


    private String getFieldValue(String excelFieldValue,T data){
        Field[] fields = ReflectUtil.getFields(clazz);
        Field realField = null;
        for(Field field : fields){
            Annotation[] annotations = field.getAnnotations();
            for(Annotation annotation : annotations){
                if(annotation.annotationType() == ExcelProperty.class){
                    ExcelProperty excelProperty = (ExcelProperty)annotation;
                    String value = excelProperty.value();
                    if(value.equals(excelFieldValue)){
                        realField = field;
                    }
                }
            }
        }
        if(realField == null){
            return null;
        }

        return String.valueOf(ReflectUtil.getFieldValue(data,realField));

    }


    private  String getCellValue(Cell cell){
        CellValue evaluate = evaluator.evaluate(cell);
        if(evaluate != null){
            CellType cellTypeEnum = evaluate.getCellTypeEnum();
            switch (cellTypeEnum){
                case STRING:
                    return evaluate.getStringValue();
                case BOOLEAN:
                    return evaluate.getBooleanValue()+"";
                case NUMERIC:
                    if(DateUtil.isCellDateFormatted(cell)){
                        Date date = DateUtil.getJavaDate(evaluate.getNumberValue());
                        return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault()).format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
                    }
                    return String.valueOf(evaluate.getNumberValue());
                case FORMULA:
                    System.out.println("-------------");
                default:
                    return null;
            }
        }
        return null;
    }


}
