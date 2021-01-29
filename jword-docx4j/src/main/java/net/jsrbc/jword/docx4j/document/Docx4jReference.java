package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.CaptionLabel;
import net.jsrbc.jword.core.document.Reference;
import org.docx4j.wml.R;

import java.util.ArrayList;
import java.util.List;

import static net.jsrbc.jword.docx4j.factory.Docx4jFactory.*;

/**
 * DOCX4J交叉引用
 * @author ZZZ on 2021/1/29 9:29
 * @version 1.0
 */
public class Docx4jReference implements Reference {

    /** 交叉引用容器 */
    private final List<R> refRuns = new ArrayList<>();

    /** {@inheritDoc} */
    @Override
    public void referTo(CaptionLabel captionLabel) {
        if (!(captionLabel instanceof Docx4jCaptionLabel))
            throw new IllegalArgumentException("captionLabel is not Docx4jCaptionLabel");
        Docx4jCaptionLabel label = (Docx4jCaptionLabel) captionLabel;
        // 清空
        this.refRuns.clear();
        // 填充引用内容
        this.refRuns.add(createFldBegin());
        this.refRuns.add(createInstrText(String.format("REF %s \\h", label.getBookmarkName())));
        this.refRuns.add(createFldSeparate());
        this.refRuns.add(createText(label.getLabel()));
        // 有章节号时，加入引用部分
        if (label.getChapterNumber() != null) {
            this.refRuns.add(createText(label.getChapterNumber()));
            this.refRuns.add(createNoBreakHyphen());
        }
        // 加入序号
        this.refRuns.add(createText(String.valueOf(label.getSequence())));
        this.refRuns.add(createFldEnd());
    }

    /**
     * 获取交叉引用DOCX4J对象
     * @return 交叉引用对应的标签集合
     */
    public List<R> getReferenceOfDocx4j() {
        return this.refRuns;
    }
}
