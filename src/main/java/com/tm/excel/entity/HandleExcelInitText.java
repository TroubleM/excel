package com.tm.excel.entity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:22
 * @Description Excel操作动态正文
 **/
public class HandleExcelInitText {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:23
     * @Description 正文单元格集合
     */
    private CellStyle textCellStyle;

    public CellStyle getTextCellStyle() {
        return textCellStyle;
    }

    public void setTextCellStyle(CellStyle textCellStyle) {
        this.textCellStyle = textCellStyle;
    }
}
