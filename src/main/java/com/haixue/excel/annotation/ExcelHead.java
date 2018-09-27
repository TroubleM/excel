package com.haixue.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 11:38
 * @Description 标题头部注解，重点注意单元格的合并使用
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelHead {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:02
     * @Description 正标题名
     */
    String name();

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:02
     * @Description 头标题单元格合并
     */
    int[] rangeAddress() default {};

}
