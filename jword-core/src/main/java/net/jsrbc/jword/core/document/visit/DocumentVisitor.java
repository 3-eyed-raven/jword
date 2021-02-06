package net.jsrbc.jword.core.document.visit;

import net.jsrbc.jword.core.document.Text;

/**
 * 文档访问者
 * @author ZZZ on 2021/2/6 15:56
 * @version 1.0
 */
public interface DocumentVisitor {
    /**
     * 访问文档
     * @param text 文档内容
     * @return 文档访问结果
     */
    default DocumentVisitResult visitText(Text text) {
        return DocumentVisitResult.CONTINUE;
    }
}
