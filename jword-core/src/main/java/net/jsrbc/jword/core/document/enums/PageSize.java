package net.jsrbc.jword.core.document.enums;

/**
 * 页面尺寸分类
 * @author ZZZ on 2021/1/29 20:21
 * @version 1.0
 */
public enum PageSize {

    /** A3纸张 */
    A3(16838, 23811, 8),

    /** A4纸张 */
    A4(11906, 16838, 9),

    /** A5纸张 */
    A5(8391, 11906, 11)
    ;

    /** 纸张宽度，单位：Twips */
    private final int width;

    /** 纸张高度，单位：Twips */
    private final int height;

    /** 尺寸编码 */
    private final int code;

    PageSize(int width, int height, int code) {
        this.width = width;
        this.height = height;
        this.code = code;
    }

    /**
     * 获取纸张宽度，单位：Twips
     * @return 获取纸张宽度
     */
    public int getWidth() {
        return width;
    }

    /**
     * 获取纸张高度，单位：Twips
     * @return 获取纸张高度
     */
    public int getHeight() {
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
