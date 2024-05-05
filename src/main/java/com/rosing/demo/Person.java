package com.rosing.demo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class Person {

    @ExcelProperty(value = "姓名",index = 0)
    private String name;
    @ExcelProperty(value="年龄",index = 1)
    private int age;
    @ExcelProperty(value="性别",index = 2)
    private String sex;
    @ExcelProperty(value="出生年月",index = 3)
    private String birthday;
    @ExcelProperty(value="出行方式",index = 4)
    private String type;
    @ExcelProperty(value="总计",index = 5)
    private String total;

}
