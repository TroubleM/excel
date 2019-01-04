package com.tm.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @auther: zhangyi
 * @date: 2019/1/4
 * @Description: excel读取相关操作注解
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelReadBean{

    /**
     * @auther: zhangyi
     * @date: 2019/1/4
     * @Description: sheet中是否存在标题
     */
    boolean hasHeaderTitle() default false;

    /**
     * @auther: zhangyi
     * @date: 2019/1/4
     * @Description: 标题所占的单元格高度,默认为0
     */
    int headerTitleHeight() default 0;

}
