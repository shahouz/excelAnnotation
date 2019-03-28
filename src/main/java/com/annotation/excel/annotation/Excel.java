package com.annotation.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel注解
 *
 * @author : tdl
 * @date : 2019/3/28 下午1:57
 **/
@Target({ElementType.FIELD, ElementType.TYPE, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    String title();

    String valueType() default "string";

    int width() default 5000;
}