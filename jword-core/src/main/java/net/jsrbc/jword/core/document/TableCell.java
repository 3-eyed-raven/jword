package net.jsrbc.jword.core.document;

import net.jsrbc.jword.core.document.enums.TableWidthType;
import net.jsrbc.jword.core.document.enums.VerticalAlignType;
import net.jsrbc.jword.core.document.enums.VerticalMergeType;

/**
 * 表格单元格
 * @author ZZZ on 2021/2/1 21:05
 * @version 1.0
 */
public interface TableCell {

    /**
     * 设定单元格宽度
     * @param width 单元格宽度
     * @param widthType 宽度类型
     */
    void setCellWidth(int width, TableWidthType widthType);

    /**
     * 设置单元格行级合并
     * @param span 合并的网格数
     */
    void setGridSpan(int span);

    /**
     * 设置垂直合并类型
     * @param verticalMergeType 垂直合并类型
     */
    void setVerticalMergeType(VerticalMergeType verticalMergeType);

    /**
     * 设置垂直对齐方式
     * @param verticalAlignType 垂直对齐方式
     */
    void setVerticalAlignType(VerticalAlignType verticalAlignType);

    /**
     * 添加段落
     * @param paragraph 段落
     */
    void addParagraph(Paragraph paragraph);
}
