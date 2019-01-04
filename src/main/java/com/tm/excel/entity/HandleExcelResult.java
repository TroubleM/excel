package com.tm.excel.entity;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:23
 * @Description 操作Excel初始化返回对象
 **/
public class HandleExcelResult {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:23
     * @Description sheet对象
     */
    private HandleExcelInitSheet handleExcelInitSheet;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:23
     * @Description 头对象
     */
    private HandleExcelInitHead handleExcelInitHead;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:23
     * @Description 列标题对象
     */
    private HandleExcelInitColumn handleExcelInitColumn;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:24
     * @Description 列动态文字对象
     */
    private HandleExcelInitText handleExcelInitText;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:24
     * @Description 对象单例化
     */
    private static HandleExcelResult handleExcelResult;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:24
     * @Description Excel操作核心对象
     */
    private HSSFWorkbook hssfWorkbook;

    /**
     * @Description 无参构造私有化
     * @Author TroubleMan
     * @date 2018/6/25 14:24
     * @param
     * @return
     **/
    private HandleExcelResult() {

    }

    /**
     * @Description 类初始化即初始化静态单例变量
     * @Author TroubleMan
     * @date 2018/6/25 14:25
     * @param
     * @return
     **/
    static {
        handleExcelResult = new HandleExcelResult();
    }


    public HandleExcelInitText getHandleExcelInitText() {
        return handleExcelInitText;
    }

    public void setHandleExcelInitText(HandleExcelInitText handleExcelInitText) {
        this.handleExcelInitText = handleExcelInitText;
    }

    public HSSFWorkbook getHssfWorkbook() {
        return hssfWorkbook;
    }

    public void setHssfWorkbook(HSSFWorkbook hssfWorkbook) {
        this.hssfWorkbook = hssfWorkbook;
    }

    public static HandleExcelResult getInstance() {
        return handleExcelResult;
    }

    public HandleExcelInitSheet getHandleExcelInitSheet() {
        return handleExcelInitSheet;
    }

    public void setHandleExcelInitSheet(HandleExcelInitSheet handleExcelInitSheet) {
        this.handleExcelInitSheet = handleExcelInitSheet;
    }

    public HandleExcelInitHead getHandleExcelInitHead() {
        return handleExcelInitHead;
    }

    public void setHandleExcelInitHead(HandleExcelInitHead handleExcelInitHead) {
        this.handleExcelInitHead = handleExcelInitHead;
    }

    public HandleExcelInitColumn getHandleExcelInitColumn() {
        return handleExcelInitColumn;
    }

    public void setHandleExcelInitColumn(HandleExcelInitColumn handleExcelInitColumn) {
        this.handleExcelInitColumn = handleExcelInitColumn;
    }

}
