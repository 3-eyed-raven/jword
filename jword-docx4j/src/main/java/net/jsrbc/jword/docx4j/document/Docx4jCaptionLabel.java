package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.CaptionLabel;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static net.jsrbc.jword.docx4j.factory.Docx4jFactory.*;

/**
 * DOCX4J题注
 * @author ZZZ on 2021/1/28 16:36
 * @version 1.0
 */
public class Docx4jCaptionLabel implements CaptionLabel {

    /** 书签ID */
    private final long bookmarkId;

    /** 书签名 */
    private final String bookmarkName;

    /** WML对象工厂 */
    private final ObjectFactory factory = Context.getWmlObjectFactory();

    /** 题注标签 */
    private final R labelRun = factory.createR();

    /** 包含章节号 */
    private final List<R> chapterRuns = new ArrayList<>();

    /** 编号方式 */
    private final List<R> sequenceRuns = new ArrayList<>();

    /**
     * 创建题注
     * @param bookmarkId 题注ID
     * @param bookmarkName 题注名
     * @return 题注
     */
    public static Docx4jCaptionLabel create(long bookmarkId, String bookmarkName) {
        return new Docx4jCaptionLabel(bookmarkId, bookmarkName);
    }

    /** {@inheritDoc} */
    @Override
    public void setLabel(String label) {
        // 先清除题注标签段
        labelRun.getContent().clear();
        // 添加题注文字
        Text text = factory.createText();
        text.setValue(label);
        labelRun.getContent().add(text);
    }

    /** {@inheritDoc} */
    @Override
    public void setChapter(String chapterStyleId, String initialNumber) {
        // 先清除章节段
        chapterRuns.clear();
        // 组成章节号段
        Collections.addAll(
                chapterRuns,
                fldBegin(),
                instrText(String.format("STYLEREF %s \\s", chapterStyleId)),
                fldSeparate(),
                text(initialNumber),
                fldEnd()
        );
    }

    /** {@inheritDoc} */
    @Override
    public void setSequence(String chapterStyleId, int initialSeq) {
        // 先清除序号段
        sequenceRuns.clear();
        // 序号关联
        R seqRun = factory.createR();
        Text seqText = factory.createText();
        seqText.setSpace("preserve");
        seqText.setValue("SEQ ");
        JAXBElement<Text> seqInstr = factory.createRInstrText(seqText);
        seqRun.getContent().add(seqInstr);
        // 组成编号段
        Collections.addAll(
                sequenceRuns,
                fldBegin(),
                preserveInstrText("SEQ "),
                labelRun,
                preserveInstrText(String.format(" \\* ARABIC \\s %s", chapterStyleId)),
                preserveInstrText(" "),
                fldSeparate(),
                text(String.valueOf(initialSeq)),
                fldEnd()
        );
    }

    /**
     * 获取题注wml标签
     * @return 题注对应的标签集合
     */
    public List<Object> getCaptionLabelOfDocx4j() {
        List<Object> list = new ArrayList<>();
        // 书签开始
        CTBookmark bookmark = factory.createCTBookmark();
        bookmark.setId(BigInteger.valueOf(this.bookmarkId));
        bookmark.setName(this.bookmarkName);
        JAXBElement<CTBookmark> bookmarkStart = factory.createBodyBookmarkStart(bookmark);
        // 书签结束
        CTMarkupRange markupRange = factory.createCTMarkupRange();
        markupRange.setId(BigInteger.valueOf(this.bookmarkId));
        JAXBElement<CTMarkupRange> bookmarkEnd = factory.createBodyBookmarkEnd(markupRange);
        // 组装结果
        list.add(bookmarkStart);
        list.add(labelRun);
        list.addAll(chapterRuns);
        list.add(noBreakHyphen());
        list.addAll(sequenceRuns);
        list.add(bookmarkEnd);
        return list;
    }

    private Docx4jCaptionLabel(long bookmarkId, String bookmarkName) {
        this.bookmarkId = bookmarkId;
        this.bookmarkName = bookmarkName;
    }
}
