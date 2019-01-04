package com.tm.excel.entity;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:20
 * @Description Excel操作标题列部分
 **/
public class HandleExcelInitColumn {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:20
     * @Description 列名名称集合
     */
    private List<String> columnNames;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 列名行集合
     */
    private List<Row> columnRows;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 列名单元格集合
     */
    private List<Cell> columnCells;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 列风格对象
     */
    private HSSFCellStyle columnCellStyle;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 每行最大单元格数量
     */
    private Integer maxCells;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:21
     * @Description 赋值的单元格索引集合数组
     */
    private int[] valueCellColumnIndex;

    public HSSFCellStyle getColumnCellStyle() {
        return columnCellStyle;
    }

    public void setColumnCellStyle(HSSFCellStyle columnCellStyle) {
        this.columnCellStyle = columnCellStyle;
    }

    public List<String> getColumnNames() {
        return columnNames;
    }

    public void setColumnNames(List<String> columnNames) {
        this.columnNames = columnNames;
    }

    public List<Row> getColumnRows() {
        return columnRows;
    }

    public void setColumnRows(List<Row> columnRows) {
        this.columnRows = columnRows;
    }

    public List<Cell> getColumnCells() {
        return columnCells;
    }

    public void setColumnCells(List<Cell> columnCells) {
        this.columnCells = columnCells;
    }

    public Integer getMaxCells() {
        return maxCells;
    }

    public void setMaxCells(Integer maxCells) {
        this.maxCells = maxCells;
    }

    public int[] getValueCellColumnIndex() {
        return valueCellColumnIndex;
    }

    public void setValueCellColumnIndex(int[] valueCellColumnIndex) {
        this.valueCellColumnIndex = valueCellColumnIndex;
    }
}
