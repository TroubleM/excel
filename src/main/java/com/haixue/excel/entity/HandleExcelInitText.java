package com.haixue.excel.entity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

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
    private HSSFCellStyle textCellStyle;

    public HSSFCellStyle getTextCellStyle() {
        return textCellStyle;
    }

    public void setTextCellStyle(HSSFCellStyle textCellStyle) {
        this.textCellStyle = textCellStyle;
    }
}
