package com.rosing.excel;

import cn.hutool.core.util.ReflectUtil;
import com.rosing.demo.ExcelProperty;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * @author luoxin
 */
public class ExcelReader<T> {


    private static MergedDataPool mergedDataPool;

    private static FormulaEvaluator evaluator;

    private List<Map<Integer,String>> headList = new ArrayList<>();

    private Class<T> myClazz = null;


    public List<T> readExcel(InputStream inputStream,String excelType,Class<T> clazz) throws Exception {
        myClazz = clazz;
        Workbook workbook = null;
        int rowNum = 2;
        if(excelType.equals("xls")){
            workbook = new HSSFWorkbook(inputStream);
        }else{
            workbook = new XSSFWorkbook(inputStream);
        }
        evaluator=workbook.getCreationHelper().createFormulaEvaluator();
        Sheet sheet = workbook.getSheetAt(0);
        mergedDataPool = new MergedDataPool(sheet);
        readHead(sheet, rowNum);
        headList.forEach(System.out::println);

        List<T> ts = readContent(sheet, rowNum);
        ts.forEach(System.out::println);
        return null;
    }

    private List<T> readContent(Sheet sheet,int rowNum) throws Exception {
        int lastRowNum = sheet.getLastRowNum();
        if(lastRowNum < rowNum){
            return null;
        }
        List<T> dataList = new ArrayList<>();
        for(int i=rowNum;i<=lastRowNum;i++){
            Row row = sheet.getRow(i);
            if(row == null){
                break;
            }
            int firstCellNum = row.getFirstCellNum();
            int lastCellNum = row.getLastCellNum();
            T t = myClazz.newInstance();
            for(int j=firstCellNum;j<lastCellNum;j++){
                Cell cell = row.getCell(j);
                String headKey = headList.get(headList.size()-1).get(j);
                String cellValue = getCellValue(cell);
                fillValue(t,cellValue,headKey);
            }
            dataList.add(t);
        }
        return dataList;
    }

    private void fillValue(T t,String value,String headKey){
        Field[] fields = ReflectUtil.getFields(myClazz);
        Field realField = null;
        for(Field field : fields){
            Annotation[] annotations = field.getAnnotations();
            for(Annotation annotation : annotations){
                Class<? extends Annotation> aClass = annotation.annotationType();
                if(aClass == ExcelProperty.class){
                    ExcelProperty excelProperty = (ExcelProperty)annotation;
                    String headName = excelProperty.value();
                    if(headName.equals(headKey)){
                        realField = field;
                    }
                }
            }
        }

        if(realField != null){
            ReflectUtil.setFieldValue(t,realField,value);
        }
    }


    private  List<Map<Integer,String>> readHead(Sheet sheet,int rowNum){
        for (int i = 0; i < rowNum; i++) {
            Map<Integer,String> rowHeadMap = new HashMap<>();
            Row row = sheet.getRow(i);
            short firstCellNum = row.getFirstCellNum();
            short lastCellNum = row.getLastCellNum();
            for(int c1 = firstCellNum;c1<lastCellNum;c1++){
                Cell cell = row.getCell(c1);
                rowHeadMap.put(c1,getCellValue(cell));
            }
            headList.add(rowHeadMap);
        }
        return headList;
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
        }else{
            Cell realCell = mergedDataPool.getRealCell(cell);
            if(realCell != null){
                return getCellValue(realCell);
            }
        }
        return null;
    }
}
