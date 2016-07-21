package com.jd.excel.annotation;

import java.lang.annotation.*;

/**
 * Created by caozhifei on 2016/7/21.
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {
    /**
     * excel 列名
     * @return
     */
    String columnName();
}

