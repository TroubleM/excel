package entity;

import com.tm.excel.annotation.ExcelColumn;
import com.tm.excel.annotation.ExcelReadBean;
import com.tm.excel.base.BaseReadExcel;

import java.math.BigDecimal;

/**
 * @auther: zhangyi
 * @date: 2019/1/4
 * @Description: 读取excel数据的接收对象
 */
@ExcelReadBean(hasHeaderTitle = true,headerTitleHeight = 2)
public class ReadExcelTestBean implements BaseReadExcel {


    /**
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  交易流水号
     */
    @ExcelColumn(name = "交易流水号")
    private String tradeNo;

    /**
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  关联订单号
     */
    @ExcelColumn(name = "关联订单号")
    private String orderNo;

    /**
     * @Author TroubleMan
     * @Date 2018/5/9 11:02
     * @Description  交易金额
     */
    @ExcelColumn(name = "金额")
    private BigDecimal netPayMoney;

    /**
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  入账公司名称
     */
    @ExcelColumn(name = "入账公司")
    private String companyName;

    /**
     * @Author TroubleMan
     * @Date 2018/5/9 11:01
     * @Description  支付方式名称
     */
    @ExcelColumn(name = "支付方式")
    private String onlinePayTypeName;

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

    public BigDecimal getNetPayMoney() {
        return netPayMoney;
    }

    public void setNetPayMoney(BigDecimal netPayMoney) {
        this.netPayMoney = netPayMoney;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public String getOnlinePayTypeName() {
        return onlinePayTypeName;
    }

    public void setOnlinePayTypeName(String onlinePayTypeName) {
        this.onlinePayTypeName = onlinePayTypeName;
    }

    @Override
    public String toString() {
        return "OrderRecordingRuleLeadingInExcelVo{" +
                "tradeNo='" + tradeNo + '\'' +
                ", orderNo='" + orderNo + '\'' +
                ", netPayMoney=" + netPayMoney +
                ", companyName='" + companyName + '\'' +
                ", onlinePayTypeName='" + onlinePayTypeName + '\'' +
                '}';
    }


}
