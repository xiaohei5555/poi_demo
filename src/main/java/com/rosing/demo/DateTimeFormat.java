package com.rosing.demo;

import java.lang.annotation.*;

/**
 * @author luoxin
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface DateTimeFormat {

    String value();

}
