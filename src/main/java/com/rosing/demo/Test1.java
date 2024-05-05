package com.rosing.demo;

import com.rosing.excel.ExcelReader;

import java.io.File;
import java.io.FileInputStream;

public class Test1 {

    public static void main(String[] args)throws Exception {

        File file = new File("d:/tmp/test.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        new ExcelReader().readExcel(fileInputStream,"xlsx",Person.class);


    }

}
