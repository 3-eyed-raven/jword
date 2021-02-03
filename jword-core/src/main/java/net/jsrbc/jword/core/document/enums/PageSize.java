package net.jsrbc.jword.core.document.enums;

/**
 * 页面尺寸分类
 * @author ZZZ on 2021/1/29 20:21
 * @version 1.0
 */
public enum PageSize {

    /** A3纸张 */
    A3(29.7, 42, 8),

    /** A4纸张 */
    A4(21, 29.7, 9),

    /** A5纸张 */
    A5(14.8, 21, 11)
    ;

    /** 纸张宽度，单位：cm */
    private final double width;

    /** 纸张高度，单位：cm */
    private final double height;

    /** 尺寸编码 */
    private final int code;

    PageSize(double width, double height, int code) {
        this.width = width;
        this.height = height;
        this.code = code;
    }

    /**
     * 获取纸张宽度，单位：Twips
     * @return 获取纸张宽度
     */
    public double getWidth() {
        return width;
    }

    /**
     * 获取纸张高度，单位：Twips
     * @return 获取纸张高度
     */
    public double getHeight() {
        return height;
    }

    /**
     * 获取纸张代码
     * @return 获取纸张代码
     */
    public int getCode() {
        return code;
    }
}
