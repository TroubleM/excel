package com.tm.excel.entity.out;

import com.tm.excel.base.BaseExcel;

import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:17
 * @Description 导入excel的数据输出总对象
 **/
public class LeadingExcelResponse<T extends BaseExcel> {

    /**
     * @Description 无参构造
     * @Author TroubleMan
     * @date 2018/6/25 14:17
     * @param
     * @return
     **/

    public LeadingExcelResponse() {

    }

    /**
     * @Description 有参构造
     * @Author TroubleMan
     * @date 2018/6/25 14:18
     * @param
     * @return
     **/

    public LeadingExcelResponse(String headerTitleValue, String[] columnHeaderTitleValues) {
        this.headerTitleValue = headerTitleValue;
        this.columnHeaderTitleValues = columnHeaderTitleValues;
    }

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:18
     * @Description 头标题内容
     */
    private String headerTitleValue;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:18
     * @Description 列头部标题字符串集合
     */
    private String[] columnHeaderTitleValues;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:18
     * @Description 实际数据集合
     */
    private List<T> columnDataValues;

    public String getHeaderTitleValue() {
        return headerTitleValue;
    }

    public void setHeaderTitleValue(String headerTitleValue) {
        this.headerTitleValue = headerTitleValue;
    }

    public String[] getColumnHeaderTitleValues() {
        return columnHeaderTitleValues;
    }

    public void setColumnHeaderTitleValues(String[] columnHeaderTitleValues) {
        this.columnHeaderTitleValues = columnHeaderTitleValues;
    }

    public List<T> getColumnDataValues() {
        return columnDataValues;
    }

    public void setColumnDataValues(List<T> columnDataValues) {
        this.columnDataValues = columnDataValues;
    }
}
