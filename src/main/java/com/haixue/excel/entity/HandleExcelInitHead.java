package com.haixue.excel.entity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:21
 * @Description Excel操作标题头对象
 **/
public class HandleExcelInitHead {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 头部显示标题
     */
    private String headTitle;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 头部行对象
     */
    private List<Row> headRows;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 头部单元格对象
     */
    private List<Cell> headCells;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description 跨行垮列参数
     */
    private int[] rangeAddress;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description 头标题处风格对象
     */
    private HSSFCellStyle headCellStyle;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:22
     * @Description 最大行数
     */
    private Integer maxRows;

    public HSSFCellStyle getHeadCellStyle() {
        return headCellStyle;
    }

    public void setHeadCellStyle(HSSFCellStyle headCellStyle) {
        this.headCellStyle = headCellStyle;
    }

    public int[] getRangeAddress() {
        return rangeAddress;
    }

    public void setRangeAddress(int[] rangeAddress) {
        this.rangeAddress = rangeAddress;
    }

    public String getHeadTitle() {
        return headTitle;
    }

    public void setHeadTitle(String headTitle) {
        this.headTitle = headTitle;
    }

    public List<Row> getHeadRows() {
        return headRows;
    }

    public void setHeadRows(List<Row> headRows) {
        this.headRows = headRows;
    }

    public List<Cell> getHeadCells() {
        return headCells;
    }

    public void setHeadCells(List<Cell> headCells) {
        this.headCells = headCells;
    }

    public Integer getMaxRows() {
        return maxRows;
    }

    public void setMaxRows(Integer maxRows) {
        this.maxRows = maxRows;
    }
}
