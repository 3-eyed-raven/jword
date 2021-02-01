package net.jsrbc.jword.core.document.enums;

/**
 * 表格单元格垂直合并类型
 * @author ZZZ on 2021/2/1 21:26
 * @version 1.0
 */
public enum VerticalMergeType {

    /**
     * 垂直合并起始点
     */
    RESTART("restart"),

    /**
     * 垂直合并中间点
     */
    CONTINUE("continue")
    ;

    private final String value;

    VerticalMergeType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
