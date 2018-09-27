package com.haixue.excel.entity.temporary;

/**
 * @Author TroubleMan
 * @Date 2018/6/25 14:18
 * @Description 风格属性对象
 **/
public class StyleAttribute {

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 是否标题居中
     */
    private Boolean isCenter;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 是否水平居中
     */
    private Boolean isAlignCenter;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 是否垂直居中
     */
    private Boolean isVerticalCenter;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 是否加粗
     */
    private Boolean isBold;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 字体类型
     */
    private String fontType;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 单元格左边线是否设置
     */
    private Boolean isLeftBorder;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 单元格右边线是否设置
     */
    private Boolean isRightBorder;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 单元格上边线是否设置
     */
    private Boolean isTopBorder;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:19
     * @Description 单元格下边线是否设置
     */
    private Boolean isBottomBorder;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:20
     * @Description 背景色
     */
    private Short background;

    /**
     * @Author TroubleMan
     * @Date 2018/6/25 14:20
     * @Description 单元格的长度
     */
    private Short width;

    /**
     * @Author TroubleMan
     * @Date 2018/8/2 10:31
     * @Description 标题字体大小
     */
    private Short headFontSize;

    /**
     * @Author TroubleMan
     * @Date 2018/8/2 10:31
     * @Description 列标题字体大小
     */
    private Short columnFontSize;

    /**
     * @Author TroubleMan
     * @Date 2018/8/2 10:31
     * @Description 内容字体大小
     */
    private Short textFontSize;

    public Short getHeadFontSize() {
        return headFontSize;
    }

    public void setHeadFontSize(Short headFontSize) {
        this.headFontSize = headFontSize;
    }

    public Short getColumnFontSize() {
        return columnFontSize;
    }

    public void setColumnFontSize(Short columnFontSize) {
        this.columnFontSize = columnFontSize;
    }

    public Short getTextFontSize() {
        return textFontSize;
    }

    public void setTextFontSize(Short textFontSize) {
        this.textFontSize = textFontSize;
    }

    public Boolean getCenter() {
        return isCenter;
    }

    public void setCenter(Boolean center) {
        isCenter = center;
    }

    public Boolean getAlignCenter() {
        return isAlignCenter;
    }

    public void setAlignCenter(Boolean alignCenter) {
        isAlignCenter = alignCenter;
    }

    public Boolean getVerticalCenter() {
        return isVerticalCenter;
    }

    public void setVerticalCenter(Boolean verticalCenter) {
        isVerticalCenter = verticalCenter;
    }

    public Boolean getBold() {
        return isBold;
    }

    public void setBold(Boolean bold) {
        isBold = bold;
    }

    public String getFontType() {
        return fontType;
    }

    public void setFontType(String fontType) {
        this.fontType = fontType;
    }

    public Boolean getLeftBorder() {
        return isLeftBorder;
    }

    public void setLeftBorder(Boolean leftBorder) {
        isLeftBorder = leftBorder;
    }

    public Boolean getRightBorder() {
        return isRightBorder;
    }

    public void setRightBorder(Boolean rightBorder) {
        isRightBorder = rightBorder;
    }

    public Boolean getTopBorder() {
        return isTopBorder;
    }

    public void setTopBorder(Boolean topBorder) {
        isTopBorder = topBorder;
    }

    public Boolean getBottomBorder() {
        return isBottomBorder;
    }

    public void setBottomBorder(Boolean bottomBorder) {
        isBottomBorder = bottomBorder;
    }

    public Short getBackground() {
        return background;
    }

    public void setBackground(Short background) {
        this.background = background;
    }

    public Short getWidth() {
        return width;
    }

    public void setWidth(Short width) {
        this.width = width;
    }

}
