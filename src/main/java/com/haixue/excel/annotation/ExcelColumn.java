package com.haixue.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 11:37
 * @Description 单元格注解，控制除头部字段中每个单元格列名内容，时间变量相关格式化形式，宽高等属性设置
 **/
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:42
     * @Description 对应excel单元格的列名
     */
    String name();

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:42
     * @Description 时间格式化
     */
    String dateFormat() default "yyyy-MM-dd";

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:42
     * @Description 列以下单元格的长度
     */
    short width() default 0;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:42
     * @Description 列的排序位
     */
    int sort() default -1;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:42
     * @Description 列标题单元格合并
     */
    int[] rangeAddress() default {};

}
