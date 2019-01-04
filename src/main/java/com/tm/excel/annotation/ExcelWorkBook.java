package com.tm.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 11:36
 * @Description 开启多Sheet操作注解
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelWorkBook {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:07
     * @Description  是否开启多sheet操作服务
     */
    boolean isOpenManySheetService() default false;


}
