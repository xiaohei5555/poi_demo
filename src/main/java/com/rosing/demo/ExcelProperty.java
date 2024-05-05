package com.rosing.demo;

import java.lang.annotation.*;

/**
 * @author luoxin
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

    String value();
    int index();

}
