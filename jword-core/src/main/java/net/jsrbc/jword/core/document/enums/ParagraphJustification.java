package net.jsrbc.jword.core.document.enums;

/**
 * 段落对齐方式
 * @author ZZZ on 2021/2/1 21:36
 * @version 1.0
 */
public enum ParagraphJustification {

    /** 居中对齐 */
    CENTER("center"),

    /** 左对齐 */
    LEFT("left"),

    /** 右对齐 */
    RIGHT("right"),

    /** 两端对齐 */
    BOTH("both"),

    /** 分散对齐 */
    DISTRIBUTE("distribute")
    ;

    private final String value;

    ParagraphJustification(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
