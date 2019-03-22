package com.tm.excel.framework;


import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import com.tm.excel.annotation.ExcelColumn;
import com.tm.excel.annotation.ExcelReadBean;
import com.tm.excel.annotation.ExcelSheet;
import com.tm.excel.base.BaseExcel;
import com.tm.excel.constants.InitConstant;
import com.tm.excel.entity.out.LeadingExcelResponse;
import com.tm.excel.exception.ExcelCellException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

/**
 * @Author TroubleMan
 * @Date 2018/5/28 11:24
 * @Description Excel文件数据读取封装入对象核心类
 **/
public class ReadExcelArrays {

    /**
     * @Author TroubleMan
     * @Date 2018/6/26 18:03
     * @Description 导入excel内容初始行数
     */
    private static final Integer DEFAULT_FIRST_ROW_NUMBER = 0;

    /**
     * @Description 根据excel文件获得里面的数据值
     * @Author TroubleMan
     * @date 2018/6/26 18:04
     * @param inputStream
     * @param clazz
     * @return
     **/

    public static <T extends BaseExcel> LeadingExcelResponse writeExcelOfArray(InputStream inputStream,
            Class<T> clazz,String suffix) throws Exception {

        // 验证集合规范
        ValidateExcelHandle.validateExcelWorkBook(clazz);

        Workbook workbook = PoiApiFactory.createWorkbook(inputStream,suffix);

        Sheet sheet = workbook.getSheetAt(InitConstant.INIT_SHEET_DEFAULT_INDEX);

        LeadingExcelResponse leadingExcelResponse = new LeadingExcelResponse(
                initHeaderTitleValue(workbook, clazz), initColumnHeaderTitleValues(sheet, clazz));

        // 获取数据内容
        leadingExcelResponse.setColumnDataValues(
                initColumnDataValues(leadingExcelResponse.getColumnHeaderTitleValues(), sheet, clazz));

        return leadingExcelResponse;
    }

    /**
     * @Description 获取标题内容
     * @Author TroubleMan
     * @date 2018/5/9 17:34
     * @param
     * @return
     **/

    public static <T extends BaseExcel> String initHeaderTitleValue(Workbook workbook, Class<T> clazz) {
        // 获取ExcelSheet注解
        ExcelReadBean excelReadBean = InitExcelHandleParam.getInstance().getExcelReadBeanAnnotation(clazz);
        // 若无标题则直接返回null
        if (!excelReadBean.hasHeaderTitle()) {
            return null;
        }
        Sheet sheet = workbook.getSheetAt(InitConstant.INIT_SHEET_DEFAULT_INDEX);
        Integer headerTitleHeight = excelReadBean.headerTitleHeight();
        //标题高度小于等于0，则返回null
        if(null == headerTitleHeight || headerTitleHeight <= 0){
            return null;
        }
        Cell headerTitleCell = null;
        // 遍历寻找实际写入了标题内容的cell，默认为不为空的第一个cell
        for (int i = 0; i < headerTitleHeight; i++) {
            Row row = sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                /**
                 * 兼容poi4.0以下写法
                 */
/*
                if (null != cell && !cell.equals("") && !(cell.getCellType() == HSSFCell.CELL_TYPE_BLANK))
                { headerTitleCell = cell; break; }*/

                if (null != cell && !cell.equals("") && !(cell.getCellType() == CellType.BLANK)) {
                    headerTitleCell = cell;
                    break;
                }

            }
        }
        if (null == headerTitleCell) {
            throw new ExcelCellException("头标题标识存在，内容未找到");
        }
        return headerTitleCell.getStringCellValue();
    }

    /**
     * @Description 获取列标题内容数组
     * @Author TroubleMan
     * @date 2018/5/9 17:34
     * @param
     * @return
     **/
    public static <T extends BaseExcel> String[] initColumnHeaderTitleValues(Sheet sheet, Class<T> clazz) {
        // 定义非标题有效数据的行号
        Integer columnHeaderRowNumber = DEFAULT_FIRST_ROW_NUMBER;
        // 获取ExcelSheet注解
        ExcelReadBean excelReadBean = InitExcelHandleParam.getInstance().getExcelReadBeanAnnotation(clazz);

        //如果设置列表标题不存在，则直接返回空数组
        if(null != excelReadBean && !excelReadBean.hasColumnTitle()){
            return new String[0];
        }

        if (excelReadBean.hasHeaderTitle()) {
            columnHeaderRowNumber = columnHeaderRowNumber + excelReadBean.headerTitleHeight();
        }

        // 列标题行
        Row columnHeaderRow = sheet.getRow(columnHeaderRowNumber);
        String[] columnHeaderTitleValues = new String[columnHeaderRow.getLastCellNum()];
        for (int i = 0; i < columnHeaderRow.getLastCellNum(); i++) {
            columnHeaderTitleValues[i] = columnHeaderRow.getCell(i).getStringCellValue();
        }
        return columnHeaderTitleValues;
    }

    /**
     * @Description 获取实际数据内容集合
     * @Author TroubleMan
     * @date 2018/5/9 17:34
     * @param
     * @return
     **/
    private static <T extends BaseExcel> List<T> initColumnDataValues(String[] headCellNames, Sheet sheet,
            Class<T> clazz) throws Exception {

        // 类字段属性对象集合
        Field[] fields = duplicateRemovalFieldArrays(clazz);

        // 类字段名称集合
        String[] fieldNames = new String[fields.length];

        // 数据应该投影的与导入数据列的对应位置集合
        int[] dataCellIndexs = new int[fields.length];

        // 反射赋值参数类型对象集合
        Class[] parameterClasses = new Class[fields.length];

        //若无列标题，则依次录入数据
        if(null == headCellNames || headCellNames.length <= 0) {
            for (int i = 0; i < fields.length; i++) {
                fieldNames[i] = fields[i].getName();
                parameterClasses[i] = fields[i].getType();
                dataCellIndexs[i] = i;
            }
        }else{
            // 将excel标题字符串和ExcelColumn的name属性的值一一对应上并排序
            for (int i = 0; i < fields.length; i++) {
                fieldNames[i] = fields[i].getName();
                ExcelColumn excelColumn = fields[i].getDeclaredAnnotation(ExcelColumn.class);
                if (excelColumn != null) {
                    parameterClasses[i] = fields[i].getType();
                    for (int j = 0; j < headCellNames.length; j++) {
                        if (excelColumn.name().equals(headCellNames[j])) {
                            dataCellIndexs[i] = j;
                        }
                    }
                }
            }
        }
        // 返回实际数据集合
        return packagingAttributeValue(InitExcelHandleParam.getInstance().getExcelReadBeanAnnotation(clazz),
                sheet, clazz, dataCellIndexs, fieldNames, parameterClasses);
    }

    /**
     * @Description 导入数据field根据是否存在ExcelColumn注解去重
     * @Author TroubleMan
     * @date 2018/5/9 17:21
     * @param
     * @return
     **/
    private static <T extends BaseExcel> Field[] duplicateRemovalFieldArrays(Class<T> clazz) {
        Field[] clazzFields = clazz.getDeclaredFields();

        // 去除没有标注ExcelColumn主机的字段
        List<Field> fieldList = new ArrayList<>();
        // 标注ExcelColumn字段个数
        for (int i = 0; i < clazzFields.length; i++) {
            if (null != clazzFields[i].getAnnotation(ExcelColumn.class)) {
                fieldList.add(clazzFields[i]);
            }
        }
        Field[] fields = new Field[fieldList.size()];
        fieldList.toArray(fields);
        return fields;
    }

    /**
     * @Description 实际excel有效数据值组装入对象
     * @Author TroubleMan
     * @date 2018/5/9 17:32
     * @param
     * @return
     **/
    private static <T extends BaseExcel> List<T> packagingAttributeValue(ExcelReadBean excelReadBean, Sheet sheet,
            Class<T> clazz, int[] dataCellIndexs, String[] fieldNames, Class<T>[] parameterClasses)
            throws Exception {

        List<T> resultList = new ArrayList<>();
        for (int i = excelReadBean.headerTitleHeight() + 1; i <= sheet.getLastRowNum(); i++) {
            Row dataRow = sheet.getRow(i);
            // 对象赋值
            Object object = clazz.newInstance();
            for (int j = 0; j < dataCellIndexs.length; j++) {
                Cell dataCell = dataRow.getCell(dataCellIndexs[j]);
                Method method = clazz
                        .getDeclaredMethod("set" + fieldNames[j].replaceFirst(fieldNames[j].substring(0, 1),
                                fieldNames[j].substring(0, 1).toUpperCase()), parameterClasses[j]);
                Class<?> parameterClazz = parameterClasses[j];
                if (null != dataCell) {
                    // 现将获得的Cell里面的值设置成字符串类型
                    /**
                     * 兼容poi3.6写法
                     */
                    /*dataCell.setCellType(Cell.CELL_TYPE_STRING);*/
                    dataCell.setCellType(CellType.STRING);
                    if (parameterClazz.isAssignableFrom(BigDecimal.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object,
                                StringUtils.isNotBlank(value) ? new BigDecimal(value) : new BigDecimal("-1"));
                    } else if (parameterClazz.isAssignableFrom(Integer.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? new Integer(value) : -1);
                    } else if (parameterClazz.isAssignableFrom(Double.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? new Double(value) : -1D);
                    } else if (parameterClazz.isAssignableFrom(String.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? value.trim() : "");
                    } else if (parameterClazz.isAssignableFrom(Float.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? new Float(value) : -1F);
                    } else if (parameterClazz.isAssignableFrom(Boolean.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? new Boolean(value) : null);
                    } else if (parameterClazz.isAssignableFrom(Long.class)) {
                        String value = dataCell.getStringCellValue();
                        method.invoke(object, StringUtils.isNotBlank(value) ? new Long(value) : -1L);
                    } else {
                        throw new ExcelCellException("导入excel不支持的字段类型" + parameterClazz);
                    }
                } else {
                    if (parameterClazz.isAssignableFrom(BigDecimal.class)) {
                        method.invoke(object, new BigDecimal("-1"));
                    } else if (parameterClazz.isAssignableFrom(Integer.class)) {
                        method.invoke(object, -1);
                    } else if (parameterClazz.isAssignableFrom(Double.class)) {
                        method.invoke(object, -1D);
                    } else if (parameterClazz.isAssignableFrom(String.class)) {
                        method.invoke(object, "");
                    } else if (parameterClazz.isAssignableFrom(Float.class)) {
                        method.invoke(object, -1F);
                    } else if (parameterClazz.isAssignableFrom(Boolean.class)) {
                        method.invoke(object, null);
                    } else if (parameterClazz.isAssignableFrom(Long.class)) {
                        method.invoke(object, -1L);
                    } else {
                        throw new ExcelCellException("导入excel不支持的字段类型" + parameterClazz);
                    }
                }
            }
            T te = (T) object;
            resultList.add(te);
        }
        return resultList;
    }

}


