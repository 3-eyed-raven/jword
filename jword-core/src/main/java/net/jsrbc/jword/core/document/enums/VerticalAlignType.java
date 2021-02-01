package net.jsrbc.jword.core.document.enums;

/**
 * 垂直对齐方式
 * @author ZZZ on 2021/2/1 21:30
 * @version 1.0
 */
public enum VerticalAlignType {

    /** 靠上对齐 */
    TOP("top"),

    /** 居中对齐 */
    CENTER("center"),

    /** 靠下对齐 */
    BOTTOM("bottom")
    ;

    private final String value;

    VerticalAlignType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
