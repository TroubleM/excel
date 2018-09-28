package com.haixue.excel.framework;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.InputStream;

/**
 * @Author TroubleMan
 * @Date 2018/5/28 11:26
 * @Description poi相关的Api类工厂
 **/
public class PoiApiFactory {

    /**
     * @param
     * @return
     * @Description 返回excel的核心对象
     * @Author TroubleMan
     * @date 2018/5/9 17:35
     **/
    public static HSSFWorkbook createHssfWorkbook() {
        return new HSSFWorkbook();
    }

    /**
     * @param
     * @return
     * @Description 导入数据时创建excel核心对象
     * @Author TroubleMan
     * @date 2018/5/9 17:35
     **/
    public static HSSFWorkbook createHssfWorkbook(InputStream inputStream) throws Exception {

        return new HSSFWorkbook(new POIFSFileSystem(inputStream));
    }

    /**
     * @param
     * @return
     * @Description 创建sheet对象
     * @Author TroubleMan
     * @date 2018/9/5 14:31
     **/
    public static Sheet createSheet(HSSFWorkbook workbook, String sheetName) {
        return workbook.createSheet(sheetName);
    }

    /**
     * @param
     * @return
     * @Description 创建sheet对象
     * @Author TroubleMan
     * @date 2018/9/5 14:32
     **/
    public static Sheet createSheet(HSSFWorkbook workbook) {
        return createSheet(workbook, null);
    }

    /**
     * @param
     * @return
     * @Description 创建行对象
     * @Author TroubleMan
     * @date 2018/9/5 14:35
     **/
    public static Row createRow(Sheet sheet, int rowNum) {
        return sheet.createRow(rowNum);
    }

    /**
     * @param
     * @return
     * @Description 创建单元格对象
     * @Author TroubleMan
     * @date 2018/9/5 14:39
     **/
    public static Cell createCell(Row row, int cellNum) {
        return row.createCell(cellNum);
    }


}
