package net.jsrbc.jword.core.document;

/**
 * 交叉引用
 * @author ZZZ on 2021/1/29 9:12
 * @version 1.0
 */
public interface Reference {
    /**
     * 引用到题注
     * @param captionLabel 题注
     */
    void referTo(CaptionLabel captionLabel);
}
