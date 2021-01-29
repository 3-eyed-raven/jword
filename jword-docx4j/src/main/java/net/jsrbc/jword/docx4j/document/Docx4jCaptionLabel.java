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

    /** WML对象工厂 */
    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    /** 书签ID */
    private final long bookmarkId;

    /** 书签名 */
    private final String bookmarkName;

    /** 标签 */
    private String label;

    /** 章节起始样式ID */
    private String chapterStyleId;

    /** 章节号 */
    private String chapterNumber;

    /** 序号 */
    private Integer sequence;

    /**
     * 创建题注
     * @param bookmarkId 书签ID，文档内不能重复，从0开始
     * @param bookmarkName 书签名，文档内不能重复
     */
    public Docx4jCaptionLabel(long bookmarkId, String bookmarkName) {
        this.bookmarkId = bookmarkId;
        this.bookmarkName = bookmarkName;
    }

    /** {@inheritDoc} */
    @Override
    public void setLabel(String label) {
        this.label = label;
    }

    /** {@inheritDoc} */
    @Override
    public void setChapter(String chapterStyleId, String chapterNumber) {
        this.chapterStyleId = chapterStyleId;
        this.chapterNumber = chapterNumber;
    }

    /** {@inheritDoc} */
    @Override
    public void setSequence(int sequence) {
        this.sequence = sequence;
    }

    /**
     * 获取书签名称
     * @return 书签名称
     */
    public String getBookmarkName() {
        return this.bookmarkName;
    }

    /**
     * 获取题注标签
     * @return 题注标签
     */
    public String getLabel() {
        return label;
    }

    /**
     * 获取章节编号
     * @return 章节编号
     */
    public String getChapterNumber() {
        return chapterNumber;
    }

    /**
     * 获取序号
     * @return 序号
     */
    public Integer getSequence() {
        return sequence;
    }

    /**
     * 获取题注DOCX4J对象
     * @return 题注对应的标签集合
     */
    public List<Object> getCaptionLabelOfDocx4j() {
        List<Object> list = new ArrayList<>();
        // 组装结果
        // 0、书签开始
        CTBookmark bookmark = FACTORY.createCTBookmark();
        bookmark.setId(BigInteger.valueOf(this.bookmarkId));
        bookmark.setName(this.bookmarkName);
        JAXBElement<CTBookmark> bookmarkStart = FACTORY.createBodyBookmarkStart(bookmark);
        list.add(bookmarkStart);
        // 1、题注
        list.add(createText(this.label));
        // 2、章节编号部分
        if (this.chapterStyleId != null && this.chapterNumber != null) {
            Collections.addAll(
                    list,
                    createFldBegin(),
                    createInstrText(String.format("STYLEREF %s \\s", this.chapterStyleId)),
                    createFldSeparate(),
                    createText(this.chapterNumber),
                    createFldEnd()
            );
            // 分隔符
            list.add(createNoBreakHyphen()); // 有章节前缀再考虑加分隔符
        }
        // 3、编号部分
        Collections.addAll(
                list,
                createFldBegin(),
                createPreserveInstrText("SEQ "),
                createText(this.label),
                createPreserveInstrText(String.format(" \\* ARABIC \\s %s", this.chapterStyleId)),
                createPreserveInstrText(" "),
                createFldSeparate(),
                createText(String.valueOf(this.sequence)),
                createFldEnd()
        );
        // 4、书签结束
        CTMarkupRange markupRange = FACTORY.createCTMarkupRange();
        markupRange.setId(BigInteger.valueOf(this.bookmarkId));
        JAXBElement<CTMarkupRange> bookmarkEnd = FACTORY.createBodyBookmarkEnd(markupRange);
        list.add(bookmarkEnd);
        return list;
    }
}
