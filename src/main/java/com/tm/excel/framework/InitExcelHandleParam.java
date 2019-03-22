package com.tm.excel.framework;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import com.tm.excel.annotation.*;
import com.tm.excel.base.BaseExcel;
import com.tm.excel.constants.InitConstant;
import com.tm.excel.entity.*;

/**
 * @Author TroubleMan
 * @Date 2018/6/26 18:00
 * @Description 初始化有效的字段及属性值, 及产生对应的数据源
 **/
public class InitExcelHandleParam {

    /**
     * @Author TroubleMan
     * @Date 2018/6/26 18:00
     * @Description 单例对象
     */
    private static InitExcelHandleParam initExcelHandleParam;

    /**
     * @Description 构造方法私有化
     * @Author TroubleMan
     * @date 2018/6/26 18:00
     * @param
     * @return
     **/

    private InitExcelHandleParam() {}

    /**
     * @Description 类加载时加载单例对象一次
     * @Author TroubleMan
     * @date 2018/6/26 18:00
     * @param
     * @return
     **/
    static {
        initExcelHandleParam = new InitExcelHandleParam();
    }

    /**
     * @Description 外部获得对象的唯一途径
     * @Author TroubleMan
     * @date 2018/6/26 18:01
     * @param
     * @return
     **/

    public static InitExcelHandleParam getInstance() {
        return initExcelHandleParam;
    }

    /**
     * @Description 初始化结构的填充数据
     * @Author TroubleMan
     * @date 2018/6/25 14:48
     * @param clazz
     * @param workbook
     * @return
     **/

    public HandleExcelResult initArchitectureData(Class<? extends BaseExcel> clazz, Workbook workbook)
            throws Exception {

        // 获得操作Excel对象
        HandleExcelResult handleExcelResult = HandleExcelResult.getInstance();

        // 初始化sheet信息
        handleExcelResult.setHandleExcelInitSheet(this.initExcelSheetData(clazz, workbook));

        // 初始化头信息
        handleExcelResult.setHandleExcelInitHead(this.initExcelHeadData(clazz, workbook));

        // 初始化列信息
        handleExcelResult.setHandleExcelInitColumn(
                this.initExcelColumnData(clazz, workbook, handleExcelResult.getHandleExcelInitSheet()));

        // 初始化动态列文字信息
        handleExcelResult.setHandleExcelInitText(this.initExcelTextData(clazz, workbook));
        return handleExcelResult;
    }

    /**
     * @Description 初始化excel的结构部分
     * @Author TroubleMan
     * @date 2018/6/25 14:48
     * @param handleExcelResult
     * @return
     **/

    public void initExcelArchitecture(HandleExcelResult handleExcelResult) {

        // 头部结构初始化
        this.initExcelHead(handleExcelResult);

        // 列标题结构初始化
        this.initExcelColumn(handleExcelResult);

    }

    /**
     * @Description 获取类的所有标注了ExcelColumn私有字段
     * @Author TroubleMan
     * @date 2018/6/25 14:49
     * @param clazz
     * @return
     **/

    public static List<Field> getEffectiveFields(Class<? extends BaseExcel> clazz) {
        List<Field> resultList = new ArrayList<Field>();
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (null != field.getAnnotation(ExcelColumn.class)) {
                resultList.add(field);
            }
        }
        return resultList;
    }

    /**
     * @Description 获取ExcelBean注解
     * @Author TroubleMan
     * @date 2018/6/27 10:21
     * @param clazz
     * @return
     **/

    public ExcelBean getExcelBeanAnnotation(Class<? extends BaseExcel> clazz) {
        return clazz.getAnnotation(ExcelBean.class);
    }

    /**
     * @Description 获取ExcelHead注解
     * @Author TroubleMan
     * @date 2018/6/25 14:49
     * @param clazz
     * @return
     **/

    public ExcelHead getExcelHeadAnnotation(Class<? extends BaseExcel> clazz) {
        return clazz.getAnnotation(ExcelHead.class);
    }

    /**
     * @Description 获取ExcelSheet注解
     * @Author TroubleMan
     * @date 2018/6/25 14:49
     * @param clazz
     * @return
     **/
    public ExcelSheet getExcelSheetAnnotation(Class<? extends BaseExcel> clazz) {
        return clazz.getAnnotation(ExcelSheet.class);
    }

    /**
     * @Author zhangyi
     * @Description: 获取excelreadbean注解
     * @Date  2019/1/4
     * @Param [clazz]
     * @return java.lang.annotation.Annotation
     **/
    public ExcelReadBean getExcelReadBeanAnnotation(Class<? extends BaseExcel> clazz){
        return clazz.getAnnotation(ExcelReadBean.class);
    }

    /**
     * @Description 获取ExcelHeadStyle注解
     * @Author TroubleMan
     * @date 2018/6/25 14:52
     * @param clazz
     * @return
     **/

    public static Annotation getExcelHeadStyleAnnotation(Class<? extends BaseExcel> clazz) {
        ExcelHeadStyle excelHeadStyle = clazz.getAnnotation(ExcelHeadStyle.class);
        if (null != excelHeadStyle) {
            return excelHeadStyle;
        }
        return getExcelDefaultStyleAnnotation(clazz);
    }

    /**
     * @Description 获取ExcelColumnStyle注解
     * @Author TroubleMan
     * @date 2018/6/25 14:52
     * @param clazz
     * @return
     **/

    public static Annotation getExcelColumnStyleAnnotation(Class<? extends BaseExcel> clazz) {
        ExcelColumnStyle excelColumnStyle = clazz.getAnnotation(ExcelColumnStyle.class);
        if (null != excelColumnStyle) {
            return excelColumnStyle;
        }
        return getExcelDefaultStyleAnnotation(clazz);
    }

    /**
     * @Description 获取ExcelTextStyle注解
     * @Author TroubleMan
     * @date 2018/6/25 14:53
     * @param clazz
     * @return
     **/

    public static Annotation getExcelTextStyleAnnotation(Class<? extends BaseExcel> clazz) {
        ExcelTextStyle excelTextStyle = clazz.getAnnotation(ExcelTextStyle.class);
        if (null != excelTextStyle) {
            return excelTextStyle;
        }
        return getExcelDefaultStyleAnnotation(clazz);
    }

    /**
     * @Description 返回默认样式注解对象
     * @Author TroubleMan
     * @date 2018/5/25 15:59
     * @param clazz
     * @return
     **/
    private static Annotation getExcelDefaultStyleAnnotation(Class<? extends BaseExcel> clazz) {
        return clazz.getAnnotation(ExcelDefaultStyle.class);
    }

    /**
     * @Description 获取字节码的字段上的ExcelColumn注解
     * @Author TroubleMan
     * @date 2018/6/25 14:53
     * @param clazz
     * @return
     **/

    public List<ExcelColumn> getExcelColumnAnnotations(Class<? extends BaseExcel> clazz) {
        // 获取类所有字段
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelColumn> handleFields = new ArrayList<ExcelColumn>();
        for (Field field : fields) {
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (null != excelColumn) {
                handleFields.add(excelColumn);
            }
        }
        return handleFields;
    }

    /**
     * @Description 初始化sheet信息对象
     * @Author TroubleMan
     * @date 2018/6/25 14:53
     * @param clazz
     * @param workbook
     * @return
     **/

    private HandleExcelInitSheet initExcelSheetData(Class<? extends BaseExcel> clazz, Workbook workbook) {

        final HandleExcelInitSheet handleExcelInitSheet = new HandleExcelInitSheet();
        final String sheetName;
        final boolean displayGridLines;
        final boolean manualBreaks;
        final boolean isAutoSort;
        // 先从ExcelBean获取注解设置的值，如果没有再从ExcelSheet中获取
        ExcelBean excelBean = this.getExcelBeanAnnotation(clazz);
        if (null != excelBean) {
            sheetName = excelBean.sheetName();
            displayGridLines = excelBean.displayGridlines();
            manualBreaks = excelBean.manualBreaks();
            isAutoSort = excelBean.isAutoSort();
        } else {
            ExcelSheet excelSheet = this.getExcelSheetAnnotation(clazz);
            sheetName = excelSheet.sheetName();
            displayGridLines = excelSheet.displayGridlines();
            manualBreaks = excelSheet.manualBreaks();
            isAutoSort = excelSheet.isAutoSort();
        }
        // 创建sheet并设置名称名称
        Sheet sheet = PoiApiFactory.createSheet(workbook, sheetName);
        // handleExcelInitSheet.setSheetName(sheetName);
        sheet.setDisplayGridlines(displayGridLines);
        sheet.setAutobreaks(manualBreaks);
        handleExcelInitSheet.setSheet(sheet);
        handleExcelInitSheet.setAutoSort(isAutoSort);
        return handleExcelInitSheet;
    }

    /**
     * @Description 初始化头部信息对象
     * @Author TroubleMan
     * @date 2018/6/25 14:53
     * @param clazz
     * @param workbook
     * @return
     **/

    private <T extends BaseExcel> HandleExcelInitHead initExcelHeadData(Class<T> clazz, Workbook workbook)
            throws InvocationTargetException, IllegalAccessException, NoSuchFieldException,
            InstantiationException {

        // 是否存在标题
        boolean hasHeaderTitle;

        // 获取ExcelBean注解
        ExcelBean excelBean = this.getExcelBeanAnnotation(clazz);

        // 获取ExcelSheet注解
        ExcelSheet excelSheet = this.getExcelSheetAnnotation(clazz);

        // 获取hasHeaderTitle的值
        if (null != excelSheet) {
            hasHeaderTitle = excelSheet.hasHeaderTitle();
        } else {
            hasHeaderTitle = excelBean.hasHeaderTitle();
        }

        HandleExcelInitHead handleExcelInitHead = new HandleExcelInitHead();

        // 如果未设置存在标题，则头标题对象返回为空
        if (!hasHeaderTitle) {
            return null;
        }

        final String headName;
        int[] headRangeAddress;

        // 获取ExcelBean注解
        if (null != excelBean) {
            headName = excelBean.headName();
            headRangeAddress = excelBean.headRangeAddress();
        } else {
            ExcelHead excelHead = this.getExcelHeadAnnotation(clazz);
            headName = excelHead.name();
            headRangeAddress = excelHead.rangeAddress();
        }

        // 设置头标题文字
        handleExcelInitHead.setHeadTitle(headName);

        // 设置跨行跨列参数数组
        handleExcelInitHead.setRangeAddress(headRangeAddress);
        if (headRangeAddress.length > 0) {
            // 设置头部标题的最大行数
            handleExcelInitHead.setMaxRows(headRangeAddress[1] - headRangeAddress[0]);
        } else {
            handleExcelInitHead.setMaxRows(InitConstant.INIT_HEAD_MAX_ROWS);
        }
        // 设置头标题单元格的样式对象
        handleExcelInitHead.setHeadCellStyle(CellStyleFactory.createHeadCellStyle(workbook, clazz));

        return handleExcelInitHead;
    }

    /**
     * @Description 初始化列表头部对象
     * @Author TroubleMan
     * @date 2018/6/25 14:54
     * @param clazz
     * @param workbook
     * @param handleExcelInitSheet
     * @return
     **/

    private <T extends BaseExcel> HandleExcelInitColumn initExcelColumnData(Class<T> clazz,
            Workbook workbook, HandleExcelInitSheet handleExcelInitSheet)
            throws InvocationTargetException, InstantiationException, IllegalAccessException,
            NoSuchFieldException {

        HandleExcelInitColumn handleExcelInitColumn = new HandleExcelInitColumn();

        // 获取sheet对象
        Sheet sheet = handleExcelInitSheet.getSheet();

        // 获取类字段上面的ExcelColumn注解
        List<ExcelColumn> excelColumns = this.getExcelColumnAnnotations(clazz);
        // 若是手动排序，则将list排序
        if (!handleExcelInitSheet.getAutoSort()) {
            ValidateExcelHandle.validateColumnSort(excelColumns);
            excelColumns.sort(Comparator.comparingInt(ExcelColumn::sort));
        }

        // 获取该sheet列名字集合
        List<String> columnNames = new ArrayList<>();

        // 初始化每行最大单元格数量
        Integer maxCells = 0;

        // 初始化赋值的单元格索引集合数组
        int[] valueCellColumnIndex = new int[excelColumns.size()];
        for (int i = 0; i < excelColumns.size(); i++) {
            valueCellColumnIndex[i] = i;
        }
        for (int i = 0; i < excelColumns.size(); i++) {
            ExcelColumn excelColumn = excelColumns.get(i);
            if (excelColumn.width() > 0) {
                sheet.setColumnWidth(i, excelColumn.width());
            }
            columnNames.add(excelColumn.name());
            int[] rangeAddress = excelColumn.rangeAddress();
            if (4 == rangeAddress.length) {
                maxCells = maxCells + rangeAddress[3] - rangeAddress[2] + 1;
                sheet.addMergedRegion(new CellRangeAddress(rangeAddress[0], rangeAddress[1], rangeAddress[2],
                        rangeAddress[3]));
            } else {
                maxCells = maxCells + 1;
            }
            if (i + 1 < excelColumns.size()) {
                valueCellColumnIndex[i + 1] = maxCells;
            }
            handleExcelInitColumn.setMaxCells(maxCells);
        }
        handleExcelInitColumn.setValueCellColumnIndex(valueCellColumnIndex);
        handleExcelInitColumn.setColumnNames(columnNames);

        // 获取列名字下每个单元格的风格对象
        handleExcelInitColumn.setColumnCellStyle(CellStyleFactory.createColumnCellStyle(workbook, clazz));

        return handleExcelInitColumn;
    }

    /**
     * @Description 初始化列表动态文字对象
     * @Author TroubleMan
     * @date 2018/6/25 14:54
     * @param clazz
     * @return
     **/

    private <T extends BaseExcel> HandleExcelInitText initExcelTextData(Class<T> clazz, Workbook workbook)
            throws InvocationTargetException, InstantiationException, IllegalAccessException,
            NoSuchFieldException {
        HandleExcelInitText handleExcelInitText = new HandleExcelInitText();
        handleExcelInitText.setTextCellStyle(CellStyleFactory.createTextCellStyle(workbook, clazz));
        return handleExcelInitText;
    }

    /**
     * @Description 初始化Excel的总标题部分
     * @Author TroubleMan
     * @date 2018/6/25 14:54
     * @param handleExcelResult
     * @return
     **/

    private void initExcelHead(HandleExcelResult handleExcelResult) {

        HandleExcelInitHead handleExcelInitHead = handleExcelResult.getHandleExcelInitHead();

        // 若handleExcelInitHead为空，则不再执行
        if (null == handleExcelInitHead) {
            return;
        }

        // 获取sheet对象
        Sheet sheet = handleExcelResult.getHandleExcelInitSheet().getSheet();

        // 创建行对象
        Row headRow = PoiApiFactory.createRow(sheet, InitConstant.EXCEL_HEAD_ROW_INDEX);

        // 获取并执行单元格合并
        CellManage.getInstance()
                .cellRangeAddress(handleExcelInitHead.getRangeAddress(), sheet,
                        handleExcelInitHead.getHeadCellStyle());

        // 创建文本列列对象
        Cell headCell = PoiApiFactory.createCell(headRow, InitConstant.EXCEL_HEAD_CELL_INDEX);

        // 设置风格和文字
        if (null != handleExcelInitHead.getHeadCellStyle()) {
            headCell.setCellStyle(handleExcelInitHead.getHeadCellStyle());
        }
        headCell.setCellValue(handleExcelInitHead.getHeadTitle());
    }

    /**
     * @Description 初始化列表标题部分
     * @Author TroubleMan
     * @date 2018/6/25 14:55
     * @param handleExcelResult
     * @return
     **/

    private void initExcelColumn(HandleExcelResult handleExcelResult) {

        HandleExcelInitColumn handleExcelInitColumn = handleExcelResult.getHandleExcelInitColumn();

        // 获取风格对象
        CellStyle columnCellStyle = handleExcelInitColumn.getColumnCellStyle();

        // 获取列表头名文字集合
        List<String> columnNames = handleExcelInitColumn.getColumnNames();

        // 获取当前集合
        Sheet sheet = handleExcelResult.getHandleExcelInitSheet().getSheet();

        // 获取HandleExcelInitHead对象
        HandleExcelInitHead handleExcelInitHead = handleExcelResult.getHandleExcelInitHead();

        // 获取行对象
        Row columnRow = PoiApiFactory.createRow(sheet, InitConstant.EXCEL_COLUMN_ROW_INDEX
                + (null == handleExcelInitHead ? 0 : handleExcelInitHead.getMaxRows()));

        // 遍历创建标题行单元格数量
        for (int i = 0; i < handleExcelInitColumn.getMaxCells(); i++) {
            Cell columnCell = PoiApiFactory.createCell(columnRow, i);
            if (null != columnCellStyle) {
                columnCell.setCellStyle(columnCellStyle);
            }
        }

        int[] valueCellColumnIndex = handleExcelInitColumn.getValueCellColumnIndex();

        // 遍历添加列头标题文字
        for (int i = 0; i < valueCellColumnIndex.length; i++) {
            Cell columnCell = columnRow.getCell(valueCellColumnIndex[i]);
            columnCell.setCellValue(columnNames.get(i));
        }
    }

}
