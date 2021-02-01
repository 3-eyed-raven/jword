package net.jsrbc.jword.core.document.enums;

/**
 * 表格水平对齐方式
 * @author ZZZ on 2021/2/1 20:54
 * @version 1.0
 */
public enum TableJustification {

    /** 左对齐 */
    LEFT("left"),

    /** 居中 */
    CENTER("center"),

    /** 右对齐 */
    RIGHT("right")
    ;

    private final String value;

    TableJustification(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
