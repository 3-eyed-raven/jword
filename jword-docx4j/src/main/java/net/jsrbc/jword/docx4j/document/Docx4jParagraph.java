package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.CaptionLabel;
import net.jsrbc.jword.core.document.Paragraph;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import static net.jsrbc.jword.docx4j.factory.Docx4jFactory.*;

/**
 * DOCX4J段落
 * @author ZZZ on 2021/1/28 15:29
 * @version 1.0
 */
public class Docx4jParagraph implements Paragraph {

    private final ObjectFactory factory = Context.getWmlObjectFactory();

    private final P p = factory.createP();

    /** 创建空段落 */
    public static Docx4jParagraph empty() {
        return new Docx4jParagraph();
    }

    /**
     * 根据文字创建段落
     * @param text 文字
     * @return 段落
     */
    public static Docx4jParagraph of(String text) {
        Docx4jParagraph paragraph = new Docx4jParagraph();
        paragraph.addText(text);
        return paragraph;
    }

    /**
     * 创建带样式的段落
     * @param styleId 样式ID
     * @param text 段落文字
     * @return 段落
     */
    public static Docx4jParagraph of(String styleId, String text) {
        Docx4jParagraph paragraph = new Docx4jParagraph();
        paragraph.setStyleId(styleId);
        paragraph.addText(text);
        return paragraph;
    }

    /** {@inheritDoc} */
    @Override
    public void setStyleId(String styleId) {
        PPr pPr = factory.createPPr();
        PPrBase.PStyle style = factory.createPPrBasePStyle();
        style.setVal(styleId);
        pPr.setPStyle(style);
        this.p.setPPr(pPr);
    }

    /** {@inheritDoc} */
    @Override
    public void addText(String text) {
        this.p.getContent().add(createText(text));
    }

    /** {@inheritDoc} */
    @Override
    public void addStyledText(String styleId, String text) {
        this.p.getContent().add(createStyledText(styleId, text));
    }

    /** {@inheritDoc} */
    @Override
    public void addCaptionLabel(CaptionLabel captionLabel) {
        if (!(captionLabel instanceof Docx4jCaptionLabel))
            throw new IllegalArgumentException("captionLabel is not Docx4jCaptionLabel");
        this.p.getContent().addAll(((Docx4jCaptionLabel)captionLabel).getCaptionLabelOfDocx4j());
    }

    /** 获取段落的WML对象 */
    public P getParagraphOfDocx4j() {
        return p;
    }

    private Docx4jParagraph() {}
}
