package com.tm.excel.framework;

import com.tm.excel.annotation.ExcelColumn;
import com.tm.excel.annotation.ExcelBean;
import com.tm.excel.annotation.ExcelHead;
import com.tm.excel.annotation.ExcelSheet;
import com.tm.excel.annotation.ExcelWorkBook;
import com.tm.excel.base.BaseExcel;
import com.tm.excel.exception.ExcelCellException;
import org.apache.commons.collections.CollectionUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA. Description:检查有关Excel产出的注解 User: Administrator Date: 2017-08-22 Time:
 * 15:35
 */
public class ValidateExcelHandle {

    /**
     * 验证传入的集合长度
     *
     * @param list
     */
    public static void validateArraysSize(List<? extends BaseExcel> list) {
        if (CollectionUtils.isEmpty(list)) {
            throw new IllegalArgumentException("传入的集合为空或集合长度为0");
        }
    }

    /**
     * 检查集合实体类中必要注解是否已标注
     *
     * @param clazz
     */
    public static void validateExcelWorkBook(Class<? extends BaseExcel> clazz) {
        if (null == ExcelBean.class) {
            if (null == clazz.getAnnotation(ExcelWorkBook.class)) {
                throw new RuntimeException("请检查集合里面的对象是否标注了com.haixue.excel.annotation.ExcelWorkBook注解");
            }
            if (null == clazz.getAnnotation(ExcelHead.class)) {
                throw new RuntimeException("请检查集合里面的对象是否标注了com.haixue.excel.annotation.ExcelHead注解");
            }
            if (null == clazz.getAnnotation(ExcelSheet.class)) {
                throw new RuntimeException("请检查集合里面的对象是否标注了com.haixue.excel.annotation.ExcelSheet注解");
            }
        }
    }

    /**
     * 验证当手动设置列排序位时 排序位的规范
     *
     * @param excelColumns
     */
    public static void validateColumnSort(List<ExcelColumn> excelColumns) {

        // 收集序列号
        List<Integer> sorts = new ArrayList<>();
        for (ExcelColumn excelColumn : excelColumns) {
            sorts.add(excelColumn.sort());
        }

        // 从小到大排序
        sorts.sort((sort1, sort2) -> {
            if (sort1 > sort2) {
                return 1;
            } else if (sort1 < sort2) {
                return -1;
            } else {
                return 0;
            }
        });

        // 执行验证
        for (int i = 0; i < sorts.size(); i++) {
            if (!sorts.get(0).equals(0)) {
                throw new RuntimeException("排序位起始位置不为0，手动设置列排序必须以0开始自增长1的自然数");
            }
            if ((sorts.get(sorts.size() - 1) - sorts.get(0)) != (sorts.size() - 1)) {
                throw new RuntimeException("排序位中存在非法增加，手动设置列排序必须以0开始自增长1的自然数");
            }
        }
    }

    /**
     * 验证单元格合并数组的规范性
     *
     * @param rangeAddress
     */
    public static void validateRangeAddress(int[] rangeAddress) throws ExcelCellException {
        if (rangeAddress.length == 4) {
            for (int i = 0; i < rangeAddress.length; i++) {
                if (rangeAddress[i] < 0) {
                    throw new ExcelCellException("数组第" + i + "个元素的值必须大于等于0");
                }
            }
            if (rangeAddress[0] > rangeAddress[1]) {
                throw new ExcelCellException("合并单元格行标号起始标号不能大于结尾标号");
            }
            if (rangeAddress[2] > rangeAddress[3]) {
                throw new ExcelCellException("合并单元格列标号起始标号不能大于结尾标号");
            }
        } else {
            throw new ExcelCellException("数组长度必须为4");
        }
    }

}
