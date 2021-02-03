package net.jsrbc.jword.core.document.enums;

/**
 * 表格宽度类型
 * @author ZZZ on 2021/2/1 20:26
 * @version 1.0
 */
public enum TableWidthType {

    /** 自动调整表格宽度 */
    AUTO("auto"),

    /** 固定表格宽度，宽度单位：cm */
    DXA("dxa"),

    /** 百分比 */
    PCT("pct"),

    /** 无宽度 */
    NIL("nil")
    ;

    private final String value;

    TableWidthType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
