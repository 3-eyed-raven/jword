package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.CaptionLabel;
import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.Reference;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import static net.jsrbc.jword.docx4j.factory.Docx4jFactory.*;

/**
 * DOCX4J段落
 * @author ZZZ on 2021/1/28 15:29
 * @version 1.0
 */
public class Docx4jParagraph implements Paragraph {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    private final P p = FACTORY.createP();

    /** {@inheritDoc} */
    @Override
    public void setStyleId(String styleId) {
        PPr pPr = FACTORY.createPPr();
        PPrBase.PStyle style = FACTORY.createPPrBasePStyle();
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

    /** {@inheritDoc} */
    @Override
    public void addReference(Reference reference) {
        if (!(reference instanceof Docx4jReference))
            throw new IllegalArgumentException("reference is not Docx4jReference");
        this.p.getContent().addAll(((Docx4jReference) reference).getReferenceOfDocx4j());
    }

    /** 获取段落的DOCX4J对象 */
    public P getParagraphOfDocx4j() {
        return p;
    }
}
