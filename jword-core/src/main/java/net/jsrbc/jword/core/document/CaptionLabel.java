package net.jsrbc.jword.core.document;

/**
 * 题注
 * @author ZZZ on 2021/1/28 15:26
 * @version 1.0
 */
public interface CaptionLabel {
    /**
     * 设置题注标签
     * @param label 题注标签
     */
    void setLabel(String label);

    /**
     * 设置章节样式
     * @param chapterStyleId 章节起始样式ID
     * @param chapterNumber 章节编号，如：1.1.1
     */
    void setChapter(String chapterStyleId, String chapterNumber);

    /**
     * 设置序号
     * @param sequence 序号，如：3
     */
    void setSequence(int sequence);
}
