package com.tm.excel.exception;

/**
 * @Author TroubleMan
 * @Date 2018/9/5 10:55
 * @Description Excel操作入参异常
 **/
public class ExcelParamException extends RuntimeException {

    private static final long serialVersionUID = 2908914409184228124L;

    public ExcelParamException(String message) {
        super(message);
    }

}
