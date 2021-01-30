package net.jsrbc.jword.core.document.enums;

/**
 * 分节符类型
 * @author ZZZ on 2021/1/30 11:13
 * @version 1.0
 */
public enum SectionType {
    /** 连续分节符 */
    CONTINUOUS("continuous"),

    /** 偶数页分节符 */
    EVEN_PAGE("evenPage"),

    /** 奇数页分节符 */
    ODD_PAGE("oddPage")
    ;
    /** 值 */
    private final String value;

    SectionType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
