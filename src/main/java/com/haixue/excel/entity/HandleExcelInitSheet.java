package com.haixue.excel.entity;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:22
 * @Description sheet表格原生属性对象
 */
public class HandleExcelInitSheet {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description 表格对象信息
     */
    private Sheet sheet;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description sheet表格的名称
     */
    private String sheetName;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description 是否自动排序
     */
    private Boolean isAutoSort;

    public Boolean getAutoSort() {
        return isAutoSort;
    }

    public void setAutoSort(Boolean autoSort) {
        isAutoSort = autoSort;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

}

