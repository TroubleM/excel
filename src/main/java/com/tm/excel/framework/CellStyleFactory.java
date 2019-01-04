package com.tm.excel.framework;

import com.tm.excel.base.BaseExcel;
import com.tm.excel.constants.FontSizeTypeConstant;
import com.tm.excel.entity.temporary.StyleAttribute;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:34
 * @Description 单元格风格对象工厂
 */
public class CellStyleFactory {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:34
     * @Description 注解父类的方法名字符串
     */
    private static final String superClassMethods = "equals,toString,hashCode,annotationType";

    /**
     * @Description 获取头标题的样式对象
     * @Author TroubleMan
     * @date 2018/6/25 14:36
     * @param workBook
     * @return
     **/

    public static HSSFCellStyle createHeadCellStyle(HSSFWorkbook workBook, Class<? extends BaseExcel> clazz)
            throws InvocationTargetException, IllegalAccessException, NoSuchFieldException,
            InstantiationException {
        return initHssfCellStyle(workBook,
                initStyleAttribute(InitExcelHandleParam.getExcelHeadStyleAnnotation(clazz)),
                FontSizeTypeConstant.HEAD_FONT_SIZE);
    }

    /**
     * @Description
     * @Author TroubleMan
     * @date 2018/6/25 14:36
     * @param workBook
     * @return
     **/

    public static <T extends BaseExcel> HSSFCellStyle createColumnCellStyle(HSSFWorkbook workBook,
            Class<T> clazz) throws NoSuchFieldException, InstantiationException, IllegalAccessException,
            InvocationTargetException {
        return initHssfCellStyle(workBook,
                initStyleAttribute(InitExcelHandleParam.getExcelColumnStyleAnnotation(clazz)),
                FontSizeTypeConstant.COLUMN_FONT_SIZE);
    }

    /**
     * @Description 获取动态内容的样式对象
     * @Author TroubleMan
     * @date 2018/6/25 14:37
     * @param workBook
     * @return
     **/

    public static <T extends BaseExcel> HSSFCellStyle createTextCellStyle(HSSFWorkbook workBook,
            Class<T> clazz) throws NoSuchFieldException, InstantiationException, IllegalAccessException,
            InvocationTargetException {
        return initHssfCellStyle(workBook,
                initStyleAttribute(InitExcelHandleParam.getExcelTextStyleAnnotation(clazz)),
                FontSizeTypeConstant.TEXT_FONT_SIZE);
    }

    /**
     * @Description 初始化风格对象的属性，从注解中获取具体值
     * @Author TroubleMan
     * @date 2018/6/25 14:38
     * @param annotation
     * @return
     **/

    private static StyleAttribute initStyleAttribute(Annotation annotation) throws InvocationTargetException,
            IllegalAccessException, NoSuchFieldException, InstantiationException {
        if (null != annotation) {
            Method[] styleAnnotationMethods = annotation.getClass().getDeclaredMethods();
            Class<? extends Object> styleAttributeClazz = StyleAttribute.class;
            StyleAttribute styleAttribute = (StyleAttribute) styleAttributeClazz.newInstance();
            for (Method styleAnnotationMethod1 : styleAnnotationMethods) {
                if (!superClassMethods.contains(styleAnnotationMethod1.getName())) {
                    Field styleAttributeField =
                            styleAttributeClazz.getDeclaredField(styleAnnotationMethod1.getName());
                    styleAttributeField.setAccessible(true);
                    styleAttributeField.set(styleAttribute, styleAnnotationMethod1.invoke(annotation, null));
                }
            }
            return styleAttribute;
        }
        return null;
    }

    /**
     * @Description 根据封装了风格属性参数的对象初始化风格对象
     * @Author TroubleMan
     * @date 2018/6/25 14:38
     * @param workBook
     * @param styleAttribute
     * @return
     **/

    private static HSSFCellStyle initHssfCellStyle(HSSFWorkbook workBook, StyleAttribute styleAttribute,
            Integer fontSizeTypeConstant) {

        // 如果属性对象为空，则返回null
        if (null == styleAttribute) {
            return null;
        }
        HSSFCellStyle cellColumnStyle = workBook.createCellStyle();

        // 自动换行
        cellColumnStyle.setWrapText(styleAttribute.getCenter());

        // 水平和竖直居中
        if (styleAttribute.getAlignCenter()) {
            /**
             * 兼容Poi3.6写法
             */
            /* cellColumnStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER); */
            cellColumnStyle.setAlignment(HorizontalAlignment.CENTER);
        }
        if (styleAttribute.getVerticalCenter()) {
            /**
             * 兼容Poi3.6写法
             */
            /* cellColumnStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); */
            cellColumnStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        // 设置列表头单元格背景
        cellColumnStyle.setFillForegroundColor(styleAttribute.getBackground());
        /**
         * 兼容Poi3.6写法
         */
        /* cellColumnStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND); */
        cellColumnStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 列表头字体风格设置
        HSSFFont fontColumn = workBook.createFont();

        // 字体大小
        if (fontSizeTypeConstant.equals(FontSizeTypeConstant.HEAD_FONT_SIZE)) {
            fontColumn.setFontHeightInPoints(styleAttribute.getHeadFontSize());
        } else if (fontSizeTypeConstant.equals(FontSizeTypeConstant.COLUMN_FONT_SIZE)) {
            fontColumn.setFontHeightInPoints(styleAttribute.getColumnFontSize());
        } else if (fontSizeTypeConstant.equals(FontSizeTypeConstant.TEXT_FONT_SIZE)) {
            fontColumn.setFontHeightInPoints(styleAttribute.getTextFontSize());
        }

        // 加粗
        if (styleAttribute.getBold()) {
            /**
             * 兼容poi3.6写法
             */
            /*fontColumn.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);*/
            fontColumn.setBold(styleAttribute.getBold());
        }

        // 字体类型
        fontColumn.setFontName(styleAttribute.getFontType());

        cellColumnStyle.setFont(fontColumn);

        // 设置单元格上下左右的边线
        if (styleAttribute.getLeftBorder()) {
            /**
             * 兼容Poi3.6写法
             */
            /*cellColumnStyle.setBorderLeft(HSSFColor.BLACK.index);*/
            cellColumnStyle.setBorderLeft(BorderStyle.THIN);
        }
        if (styleAttribute.getRightBorder()) {
            /**
             * 兼容Poi3.6写法
             */
            /*cellColumnStyle.setBorderRight(HSSFColor.BLACK.index);*/
            cellColumnStyle.setBorderRight(BorderStyle.THIN);
        }
        if (styleAttribute.getTopBorder()) {
            /**
             * 兼容Poi3.6写法
             */
            /*cellColumnStyle.setBorderTop(HSSFColor.BLACK.index);*/
            cellColumnStyle.setBorderTop(BorderStyle.THIN);
        }
        if (styleAttribute.getBottomBorder()) {
            /**
             * 兼容Poi3.6写法
             */
            /*cellColumnStyle.setBorderBottom(HSSFColor.BLACK.index);*/
            cellColumnStyle.setBorderBottom(BorderStyle.THIN);
        }
        return cellColumnStyle;
    }

}
