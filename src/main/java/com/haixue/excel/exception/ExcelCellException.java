package com.haixue.excel.exception;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:25
 * @Description 单元格处理时的异常类
 */
public class ExcelCellException extends RuntimeException {

    private static final long serialVersionUID = -686165125401655258L;

    public ExcelCellException(String message) {
        super(message);
    }

}
