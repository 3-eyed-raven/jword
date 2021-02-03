package net.jsrbc.jword.core.document;

import net.jsrbc.jword.core.document.enums.TableJustification;
import net.jsrbc.jword.core.document.enums.TableWidthType;

/**
 * 表格
 * @author ZZZ on 2021/2/1 20:19
 * @version 1.0
 */
public interface Table {

    /**
     * 设置表格样式
     * @param styleId 表格样式ID
     */
    void setStyle(String styleId);

    /**
     * 设置表格宽度
     * @param width 宽度，如果宽度类型为auto，宽度则初始为0
     * @param type 宽度类型
     */
    void setWidth(double width, TableWidthType type);

    /**
     * 设置表格对齐方式
     * @param tableJustification 表格对齐方式
     */
    void setJustification(TableJustification tableJustification);

    /**
     * 添加表格行
     * @param row 表格行
     */
    void addRow(TableRow row);
}
