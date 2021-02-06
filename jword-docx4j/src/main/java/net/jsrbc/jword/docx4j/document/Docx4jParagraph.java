package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.CaptionLabel;
import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.Reference;
import net.jsrbc.jword.core.document.Section;
import net.jsrbc.jword.core.document.Text;
import net.jsrbc.jword.core.document.enums.ParagraphJustification;
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

    /** 段落对象 */
    private final P p = FACTORY.createP();

    /** {@inheritDoc} */
    @Override
    public void setStyleId(String styleId) {
        PPr pPr = getPpr();
        PPrBase.PStyle style = FACTORY.createPPrBasePStyle();
        style.setVal(styleId);
        pPr.setPStyle(style);
    }

    /** {@inheritDoc} */
    @Override
    public void setJustification(ParagraphJustification justification) {
        PPr pPr = getPpr();
        Jc jc = FACTORY.createJc();
        jc.setVal(JcEnumeration.valueOf(justification.name()));
        pPr.setJc(jc);
    }

    /** {@inheritDoc} */
    @Override
    public void setSection(Section section) {
        if (!(section instanceof Docx4jSection))
            throw new IllegalArgumentException("section is not Docx4jSection");
        getPpr().setSectPr(((Docx4jSection) section).getSectionOfDocx4j());
    }

    /** {@inheritDoc} */
    @Override
    public void addText(Text text) {
        this.p.getContent().add(((Docx4jText) text).getTextOfDocx4j());
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

    /**
     * 获取段落属性
     * @return 段落属性
     */
    private PPr getPpr() {
        if (p.getPPr() == null) {
            PPr pPr = FACTORY.createPPr();
            p.setPPr(pPr);
            return pPr;
        }
        return p.getPPr();
    }
}
