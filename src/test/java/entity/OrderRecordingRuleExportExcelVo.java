package entity;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.Date;

import com.haixue.excel.annotation.*;
import com.haixue.excel.base.BaseExcel;

/**
 * @Author TroubleMan
 * @Date 2018/5/9 10:55
 * @Description 订单入账记录导出实体
 **/
@ExcelSheet(sheetName = "导出测试数据")
@ExcelHead(name = "导出测试")
@ExcelWorkBook
public class OrderRecordingRuleExportExcelVo implements BaseExcel, Serializable {

    private static final long serialVersionUID = -160224668300165860L;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  交易流水号
     */
    @ExcelColumn(name = "交易流水号")
    private String tradeNo;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  关联订单号
     */
    @ExcelColumn(name = "关联订单号")
    private String orderNo;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  入账公司名称
     */
    @ExcelColumn(name = "入账公司")
    private String companyName;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  支付方式名称
     */
    @ExcelColumn(name = "支付方式")
    private String onlinePayTypeName;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:02
     * @Description  入账商户号
     */
    @ExcelColumn(name = "入账商户号")
    private String accountNumber;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:02
     * @Description  交易金额
     */
    @ExcelColumn(name = "金额")
    private BigDecimal netPayMoney;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:02
     * @Description  付款时间
     */
    @ExcelColumn(name = "付款时间",dateFormat = "yyyy-MM-dd HH:mm:ss")
    private Date payTime;

    /*
     * @Author TroubleMan
     * @Date 2018/5/9 11:02
     * @Description  最后操作人名称
     */
    @ExcelColumn(name = "最后操作人名称")
    private String operationName;

    public String getOnlinePayTypeName() {
        return onlinePayTypeName;
    }

    public void setOnlinePayTypeName(String onlinePayTypeName) {
        this.onlinePayTypeName = onlinePayTypeName;
    }

    public String getTradeNo() {
        return tradeNo;
    }

    public void setTradeNo(String tradeNo) {
        this.tradeNo = tradeNo;
    }

    public String getOrderNo() {
        return orderNo;
    }

    public void setOrderNo(String orderNo) {
        this.orderNo = orderNo;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getAccountNumber() {
        return accountNumber;
    }

    public void setAccountNumber(String accountNumber) {
        this.accountNumber = accountNumber;
    }

    public BigDecimal getNetPayMoney() {
        return netPayMoney;
    }

    public void setNetPayMoney(BigDecimal netPayMoney) {
        this.netPayMoney = netPayMoney;
    }

    public Date getPayTime() {
        return payTime;
    }

    public void setPayTime(Date payTime) {
        this.payTime = payTime;
    }

    public String getOperationName() {
        return operationName;
    }

    public void setOperationName(String operationName) {
        this.operationName = operationName;
    }

    @Override
    public String toString() {
        return "OrderRecordingRuleExportExcelVo{" +
                "tradeNo='" + tradeNo + '\'' +
                ", orderNo='" + orderNo + '\'' +
                ", companyName='" + companyName + '\'' +
                ", onlinePayTypeName='" + onlinePayTypeName + '\'' +
                ", accountNumber='" + accountNumber + '\'' +
                ", netPayMoney=" + netPayMoney +
                ", payTime=" + payTime +
                ", operationName='" + operationName + '\'' +
                '}';
    }
}
