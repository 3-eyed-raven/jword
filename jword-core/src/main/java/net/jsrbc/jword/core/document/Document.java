package net.jsrbc.jword.core.document;

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
     * 添加段落
     * @param paragraph 段落对象
     */
    void addParagraph(Paragraph paragraph);

    /**
     * 保存文档
     */
    void save() throws IOException;

    /**
     * 另存到目标地址
     * @param dest 目标地址
     */
    void saveAs(Path dest) throws IOException;

    /**
     * 另存到目标地址
     * @param dest 输出流
     */
    void saveAs(OutputStream dest) throws IOException;
}
