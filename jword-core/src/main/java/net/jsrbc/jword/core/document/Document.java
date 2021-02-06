package net.jsrbc.jword.core.document;

import net.jsrbc.jword.core.document.visit.DocumentVisitor;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;

/**
 * word文档对象，对应 w:document
 * @author ZZZ on 2021/1/25 20:15
 * @version 1.0
 */
public interface Document {

    /**
     * 访问文档内容
     * @param visitor 文档访问者
     */
    void walkBody(DocumentVisitor visitor);

    /**
     * 访问页眉
     * @param visitor 文档访问者
     */
    void walkHeader(DocumentVisitor visitor);

    /**
     * 访问页脚
     * @param visitor 文档访问者
     */
    void walkFooter(DocumentVisitor visitor);

    /**
     * 添加表格
     * @param table 表格
     */
    void addTable(Table table);

    /**
     * 添加段落
     * @param paragraph 段落对象
     */
    void addParagraph(Paragraph paragraph);

    /**
     * 添加图片
     * @param target 目标段落
     * @param path 图片路径
     * @param width 图片宽度，单位：cm
     * @param height 图片高度，单位：cm
     * @throws IOException 图片无法读取时报异常
     */
    void addDrawing(Paragraph target, Path path, double width, double height) throws IOException;

    /**
     * 添加分节符
     * @param section 分节符
     */
    void addSection(Section section);

    /**
     * 更新目录，请注意：不支持更新页码
     */
    void updateTableOfContent();

    /**
     * 保存文档
     * @throws IOException 文件无法打开时报异常
     */
    void save() throws IOException;

    /**
     * 另存到目标地址
     * @param dest 目标地址
     * @throws IOException 文件无法打开时报异常
     */
    void saveAs(Path dest) throws IOException;

    /**
     * 另存到目标地址
     * @param dest 输出流
     * @throws IOException 文件无法打开时报异常
     */
    void saveAs(OutputStream dest) throws IOException;
}
