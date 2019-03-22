package com.tm.excel.framework;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:26
 * @Description 单元格共同方法操作类
 **/
public class CellManage {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:28
     * @Description 对象单例化
     */
    private static CellManage cellManage;

    /**
     * @Description 构造方法私有化
     * @Author TroubleMan
     * @date 2018/6/25 14:29
     * @param
     * @return
     **/
    private CellManage() {

    }

    /**
     * @Description 类初始化时静态变量初始化
     * @Author TroubleMan
     * @date 2018/6/25 14:29
     * @param
     * @return
     **/
    static {
        cellManage = new CellManage();
    }

    /**
     * @Description 外部获取对象唯一入口
     * @Author TroubleMan
     * @date 2018/6/25 14:30
     * @param
     * @return
     **/

    public static CellManage getInstance() {
        return cellManage;
    }

    /**
     * @Author:TroubleMan
     * @Description:验证合并单元格参数
     * @Date:2018/1/26/026
     * @param:[rangeAddress, sheet, cellStyle]
     * @return:void
     */
    public void cellRangeAddress(int[] rangeAddress, Sheet sheet, CellStyle cellStyle) {
        if (rangeAddress.length > 0) {
            ValidateExcelHandle.validateRangeAddress(rangeAddress);
            // 进行每行row判断，没有创建，则创建
            for (int i = 0; i < rangeAddress[1] - rangeAddress[0] + 1; i++) {
                Row row = sheet.getRow(i);
                if (null == row) {
                    row = sheet.createRow(i);
                }

                // 进行合并单元格的每个单元格判断，没有创建，则创建
                for (int j = 0; j < rangeAddress[3] - rangeAddress[2] + 1; j++) {
                    Cell cell = row.getCell(j);
                    if (null == cell) {
                        cell = row.createCell(j);
                        if (null != cellStyle) {
                            cell.setCellStyle(cellStyle);
                        }
                    }
                }
            }
            sheet.addMergedRegion(
                    new CellRangeAddress(rangeAddress[0], rangeAddress[1], rangeAddress[2], rangeAddress[3]));
        }
    }

}
