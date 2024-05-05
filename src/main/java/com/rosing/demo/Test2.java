package com.rosing.demo;

import com.rosing.excel.ExcelReader;
import com.rosing.excel.ExcelWriter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Test2 {

    public static void main(String[] args)throws Exception {

        File file = new File("d:/tmp/test.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        new ExcelWriter(Person.class).doWrite(getPersons(),getHead(),fos);
    }


    private static List<Person> getPersons(){
        List<Person> personList = new ArrayList<>();
        personList.add(new Person("zhangsan",12,"nan","1989-03-02","car","cccc"));
        personList.add(new Person("zhangsan",13,"nv","1989-03-03","way","cccc"));
        personList.add(new Person("zhangsan",14,"nv","1989-03-02","car","cccc"));
        personList.add(new Person("wangwu",15,"nan","1989-03-02","wa","bbbb"));
        personList.add(new Person("lisi",15,"nan","1989-03-02","car","bbbb"));
        return personList;
    }

    private static List<List<String>> getHead(){
        List<List<String>> headList = new ArrayList<>();
        headList.add(Arrays.asList("information","information","information","specify","specify","specify"));
        headList.add(Arrays.asList("姓名","年龄","性别","出生年月","出行方式","总计"));
        return headList;
    }



}
