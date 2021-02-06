package net.jsrbc.jword.core.document;

/**
 * 文本
 * @author ZZZ on 2021/2/6 16:03
 * @version 1.0
 */
public interface Text {

    /**
     * 设置文本样式
     * @param styleId 文本样式ID
     */
    void setStyleId(String styleId);

    /**
     * 设置文本内容
     * @param value 文本内容
     */
    void setValue(String value);

    /**
     * 获取文本内容
     */
    String getValue();
}
