package net.jsrbc.jword.core.document;

import net.jsrbc.jword.core.document.enums.ParagraphJustification;

import java.nio.file.Path;

/**
 * 段落对象
 * @author ZZZ on 2021/1/28 15:17
 * @version 1.0
 */
public interface Paragraph {

    /**
     * 设置整个段落样式
     * @param styleId 段落样式ID
     */
    void setStyleId(String styleId);

    /**
     * 设置段落对齐方式
     * @param justification 对齐方式
     */
    void setJustification(ParagraphJustification justification);

    /**
     * 设置分节符
     * @param section 分节符
     */
    void setSection(Section section);

    /**
     * 追加段落文字
     * @param text 段落文字
     */
    void addText(String text);

    /**
     * 追加带样式的段落文字
     * @param styleId 段落ID
     * @param text 段落文字
     */
    void addStyledText(String styleId, String text);

    /**
     * 追加题注
     * @param captionLabel 题注
     */
    void addCaptionLabel(CaptionLabel captionLabel);

    /**
     * 追加交叉引用
     * @param reference 交叉引用
     */
    void addReference(Reference reference);
}
