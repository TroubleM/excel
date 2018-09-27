package com.haixue.excel.annotation;

import com.haixue.excel.constants.CellBackgroundConstant;
import com.haixue.excel.constants.FontTypeConstant;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 11:38
 * @Description 单元格风格注解，控制内容头以及内容字体的样式
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumnStyle {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 是否标题居中
     */
    boolean isCenter() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 是否水平居中
     */
    boolean isAlignCenter() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 是否垂直居中
     */
    boolean isVerticalCenter() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 字体大小
     */
    short columnFontSize() default 12;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 是否加粗
     */
    boolean isBold() default false;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:43
     * @Description 字体类型
     */
    String fontType() default FontTypeConstant.BLACK_TYPEFACE;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 单元格左边线是否设置
     */
    boolean isLeftBorder() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 单元格右边线是否设置
     */
    boolean isRightBorder() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 单元格上边线是否设置
     */
    boolean isTopBorder() default true;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 单元格下边线是否设置
     */
    boolean isBottomBorder() default true;



    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 背景色
     */
    short background() default CellBackgroundConstant.WHITE;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 11:44
     * @Description 单元格的长度
     */
    short width() default 0;
}
