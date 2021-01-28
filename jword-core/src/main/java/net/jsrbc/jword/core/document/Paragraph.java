package net.jsrbc.jword.core.document;

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
}
