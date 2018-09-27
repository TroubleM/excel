package com.haixue.excel.framework;

import com.haixue.excel.annotation.ExcelColumn;
import com.haixue.excel.base.BaseExcel;
import com.haixue.excel.entity.HandleExcelResult;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @Author TroubleMan
 * @Date 2018/5/9 17:35
 * @Description 生产excel输入流的核心
 **/
public class ProduceExcelInputStream {

    /**
     * @Author TroubleMan
     * @Date 2018/6/26 18:02
     * @Description 初始化执行对象的创建
     */
    private static InitExcelHandleParam initExcelHandleParam = InitExcelHandleParam.getInstance();

    /**
     * 产生excel的输入流对象
     *
     * @param list
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T extends BaseExcel> InputStream createExcelInputStream(List<T> list,
                                                                           Class<? extends BaseExcel> clazz) throws Exception {
        // 验证集合规范
        ValidateExcelHandle.validateExcelWorkBook(clazz);
        return produceInputStream(initFieldValues(list, initExcelArchitecture(clazz), clazz));
    }

    /**
     * @Description 返回excel核心对象，供拓展开发
     * @Author TroubleMan
     * @date 2018/5/11 14:30
     * @param
     * @return
     **/
    public static <T extends BaseExcel> HSSFWorkbook createExcelHssfWorkbook(List<T> list,
                                                                             Class<? extends BaseExcel> clazz) throws Exception {
        // 验证集合规范
        ValidateExcelHandle.validateExcelWorkBook(clazz);
        return initFieldValues(list, initExcelArchitecture(clazz), clazz);
    }

    /**
     * @Description 初始化Excel表结构
     * @Author TroubleMan
     * @date 2018/6/26 18:02
     * @param clazz
     * @return
     **/
    public static HandleExcelResult initExcelArchitecture(Class<? extends BaseExcel> clazz) throws Exception {

        // 创建excel的核心对象
        HSSFWorkbook workbook = PoiApiFactory.createHssfWorkbook();

        // 初始化结构数据
        HandleExcelResult handleExcelResult = initExcelHandleParam.initArchitectureData(clazz, workbook);

        // 初始化结构
        initExcelHandleParam.initExcelArchitecture(handleExcelResult);

        // 设置excel核心对象
        handleExcelResult.setHssfWorkbook(workbook);

        return handleExcelResult;
    }

    /**
     * @Description 返回excel文件的输入流
     * @Author TroubleMan
     * @date 2018/6/26 18:03
     * @param workbook
     * @return
     **/
    public static InputStream produceInputStream(HSSFWorkbook workbook) throws Exception {

        // 有关返回参数
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            workbook.write(outputStream);
        } finally {
            outputStream.close();
        }
        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    /**
     * @Description 动态内容拼装
     * @Author TroubleMan
     * @date 2018/6/26 18:02
     * @param list
     * @param handleExcelResult
     * @param clazz
     * @return
     **/
    private static <T extends BaseExcel> HSSFWorkbook initFieldValues(List<T> list,
                                                                      HandleExcelResult handleExcelResult, Class<? extends BaseExcel> clazz) throws Exception {

        // 获取风格对象
        HSSFCellStyle textCellStyle = handleExcelResult.getHandleExcelInitText().getTextCellStyle();

        // 获取标注有ExcelColumn注解的字段
        List<Field> fields = InitExcelHandleParam.getEffectiveFields(clazz);

        // 获取sheet，此处暂时只默认一个sheet
        Sheet sheet = handleExcelResult.getHandleExcelInitSheet().getSheet();

        // 遍历集合为动态数据行
        Integer sheetLastRowNum = sheet.getLastRowNum();

        // 每行单元格数量
        Integer maxCells = handleExcelResult.getHandleExcelInitColumn().getMaxCells();

        // 每行需要设置值的Cell的索引集合
        int[] valueCellColumnIndex = handleExcelResult.getHandleExcelInitColumn().getValueCellColumnIndex();

        for (int i = 0; i < list.size(); i++) {
            Row valueRow = PoiApiFactory.createRow(sheet, i + sheetLastRowNum + 1);
            for (int c = 0; c < maxCells; c++) {
                Cell valueCell = PoiApiFactory.createCell(valueRow, c);
                if (null != textCellStyle) {
                    valueCell.setCellStyle(textCellStyle);
                }
            }
            // 先将需要的单元格Cell全部创建好
            // 遍历字段集合为每行的动态数据列
            if (!handleExcelResult.getHandleExcelInitSheet().getAutoSort()) {
                for (int j = 0; j < fields.size(); j++) {
                    ExcelColumn excelColumn = fields.get(j).getAnnotation(ExcelColumn.class);
                    for (int c = 0; c < maxCells; c++) {
                        Cell valueCell = PoiApiFactory.createCell(valueRow, excelColumn.sort());
                        if (null != textCellStyle) {
                            valueCell.setCellStyle(textCellStyle);
                        }
                    }
                    Cell valueCell = valueRow.getCell(valueCellColumnIndex[j]);
                    fields.get(j).setAccessible(true);
                    Object object = fields.get(j).get(list.get(i));
                    if (object instanceof Date) {
                        valueCell.setCellValue(
                                new SimpleDateFormat(excelColumn.dateFormat()).format((Date) object));
                    } else {
                        valueCell.setCellValue(object.toString());
                    }
                    if (4 == excelColumn.rangeAddress().length) {
                        sheet.addMergedRegion(
                                new CellRangeAddress(excelColumn.rangeAddress()[0] + sheetLastRowNum + 1,
                                        excelColumn.rangeAddress()[1] + sheetLastRowNum + 1,
                                        excelColumn.rangeAddress()[2], excelColumn.rangeAddress()[3]));
                    }
                }
            } else {
                for (int j = 0; j < fields.size(); j++) {
                    Field field = fields.get(j);
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    Cell valueCell = valueRow.getCell(valueCellColumnIndex[j]);
                    field.setAccessible(true);
                    Object object = field.get(list.get(i));
                    if (null == object) {
                        valueCell.setCellValue("");
                    } else {
                        if (object instanceof Date) {
                            valueCell.setCellValue(
                                    new SimpleDateFormat(field.getAnnotation(ExcelColumn.class).dateFormat())
                                            .format((Date) object));
                        } else {
                            valueCell.setCellValue(object.toString());
                        }
                    }
                    if (4 == excelColumn.rangeAddress().length) {
                        sheet.addMergedRegion(new CellRangeAddress(
                                excelColumn.rangeAddress()[0] + sheet.getLastRowNum() - 1,
                                excelColumn.rangeAddress()[1] + sheet.getLastRowNum() - 1,
                                excelColumn.rangeAddress()[2], excelColumn.rangeAddress()[3]));
                    }
                }
            }
        }
        return handleExcelResult.getHssfWorkbook();
    }

}


