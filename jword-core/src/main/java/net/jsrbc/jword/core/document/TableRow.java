package net.jsrbc.jword.core.document;

/**
 * 表格行
 * @author ZZZ on 2021/2/1 21:04
 * @version 1.0
 */
public interface TableRow {

    /**
     * 添加表格单元格
     * @param cell 单元格
     */
    void addCell(TableCell cell);
}
