package com.tm.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 11:37
 * @Description Excel操作相关Bean默认注解
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelBean {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description sheet的名称
     */
    String sheetName();

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description 是否保留单元格的网格线条
     */
    boolean displayGridlines() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description 是否开启手动设置列宽
     */
    boolean manualBreaks() default false;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description sheet的列是否默认排序
     */
    boolean isAutoSort() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description 此sheet中是否存在标题
     */
    boolean hasHeaderTitle() default false;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description 标题默认所占的单元格高度
     */
    int headerTitleHeight() default 0;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:02
     * @Description 正标题名
     */
    String headName();

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:02
     * @Description 头标题单元格合并
     */
    int[] headRangeAddress() default {};

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:05
     * @Description 初始化打开时的网格比例大小
     */
    int zoom() default 120;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:07
     * @Description 是否开启多sheet操作服务
     */
    boolean isOpenManySheetService() default false;

}
