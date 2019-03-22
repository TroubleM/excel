package com.tm.excel.framework;

import com.tm.excel.base.BaseExcel;
import com.tm.excel.entity.out.LeadingExcelResponse;
import com.tm.excel.exception.ExcelParamException;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2017-08-22 Time: 15:37
 * @Description 实际产出Excel的类， 可产出不同io形式的Excel结果
 **/
public class ExcelFactory {

    /**
     * @Description 产生输入流形式的excel文件信息,导出框架的统一API调用入口， 产出InputStream形式结果
     * @Author TroubleMan
     * @date 2018/6/25 14:40
     * @param list
     * @param clazz
     * @throws Exception
     **/
    public static <T extends BaseExcel> InputStream produceExcelOfInputStream(List<T> list, Class<T> clazz,
                                                                              String suffix)
            throws Exception {
        return ProduceExcelInputStream.createExcelInputStream(list, clazz, suffix);
    }

    /**
     * @Description 产生输入流形式的excel文件信息,导出框架的统一API调用入口，产出InputStream形式结果
     * @Author TroubleMan
     * @date 2018/9/5 10:53
     * @param list
     **/
    public static <T extends BaseExcel> InputStream produceExcelOfInputStream(List<T> list,String suffix) throws Exception {
        if (CollectionUtils.isEmpty(list)) {
            throw new ExcelParamException("导出Excel时集合为空或者长度为0");
        }
        return ProduceExcelInputStream.createExcelInputStream(list, list.get(0).getClass(),suffix);
    }

    /**
     * @Description 产生输入流形式的excel文件信息,导出框架的统一API调用入口，产出HSSFWorkbook核心api对象
     * @Author TroubleMan
     * @date 2018/5/11 14:29
     * @param list
     * @param clazz
     **/
    public static <T extends BaseExcel> Workbook produceExcelOfHssfWorkbook(List<T> list, Class<T> clazz,
                                                                            String suffix)
            throws Exception {
        return ProduceExcelInputStream.createExcelWorkbook(list, clazz,suffix);
    }

    /**
     * @Description 产生输入流形式的excel文件信息,导出框架的统一API调用入口，产出HSSFWorkbook核心api对象
     * @Author TroubleMan
     * @date 2018/5/11 14:29
     * @param list
     **/
    public static <T extends BaseExcel> Workbook produceExcelOfHssfWorkbook(List<T> list,
                                                                            String suffix)
            throws Exception {
        if (CollectionUtils.isEmpty(list)) {
            throw new ExcelParamException("导出Excel时集合为空或者长度为0");
        }
        return ProduceExcelInputStream.createExcelWorkbook(list, list.get(0).getClass(),suffix);
    }

    /**
     * @Description 将文件流数据写入并返回相关Excel信息对象
     * @Author TroubleMan
     * @date 2018/5/9 15:19
     * @param inputStream
     * @param clazz
     **/
    public static <T extends BaseExcel> LeadingExcelResponse<T> writeExcelOfInputStream(
            InputStream inputStream, Class<T> clazz,String suffix) throws Exception {
        return ReadExcelArrays.writeExcelOfArray(inputStream, clazz, suffix);
    }

}
