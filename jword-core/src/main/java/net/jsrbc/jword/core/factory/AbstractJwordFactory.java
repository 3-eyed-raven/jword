package net.jsrbc.jword.core.factory;

import net.jsrbc.jword.core.document.*;

import java.io.IOException;
import java.nio.file.Path;

/**
 * Jword的抽象工厂
 * @author ZZZ on 2021/2/4 9:49
 * @version 1.0
 */
public interface AbstractJwordFactory {

    /**
     * 创建文档对象
     * @param path 模板文件路径
     * @return 文档
     * @throws IOException 文件不存在或无法访问时抛出异常
     */
    Document loadDocument(Path path) throws IOException;

    /**
     * 创建段落
     * @return 段落
     */
    Paragraph createParagraph();

    /**
     * 创建题注
     * @return 题注
     */
    CaptionLabel createCaptionLabel();

    /**
     * 创建交叉引用
     * @return 交叉引用
     */
    Reference createReference();

    /**
     * 创建分节符
     * @return 分节符
     */
    Section createSection();

    /**
     * 创建表格
     * @return 表格
     */
    Table createTable();

    /**
     * 创建表格行
     * @return 表格行
     */
    TableRow createTableRow();

    /**
     * 创建表格单元格
     * @return 单元格
     */
    TableCell createTableCell();
}
