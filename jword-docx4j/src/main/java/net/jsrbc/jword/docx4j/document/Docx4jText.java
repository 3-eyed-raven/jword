package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Text;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RStyle;

import javax.xml.bind.JAXBElement;

/**
 * Docx4j文本
 * @author ZZZ on 2021/2/6 16:05
 * @version 1.0
 */
public class Docx4jText implements Text {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    /** 运行对象 */
    private final R r;

    /** 文本内容 */
    private String text;

    /** {@inheritDoc} */
    @Override
    public void setStyleId(String styleId) {
        RPr rPr = getOrCreateRpr();
        RStyle style = FACTORY.createRStyle();
        style.setVal(styleId);
        rPr.setRStyle(style);
    }

    /** {@inheritDoc} */
    @Override
    @SuppressWarnings("unchecked")
    public void setValue(String text) {
        this.text = text;
        if (r.getContent().isEmpty()) {
            org.docx4j.wml.Text t = FACTORY.createText();
            t.setValue(text);
            t.setSpace("preserve");
            this.r.getContent().add(t);
        } else {
            for (Object content : r.getContent()) {
                if (content instanceof JAXBElement && ((JAXBElement<?>) content).getValue() instanceof org.docx4j.wml.Text) {
                    org.docx4j.wml.Text t = ((JAXBElement<org.docx4j.wml.Text>) content).getValue();
                    t.setValue(text);
                }
            }
        }
    }

    /** {@inheritDoc} */
    @Override
    public String getValue() {
        return this.text;
    }

    /**
     * 获取运行对象
     * @return 运行对象
     */
    public R getTextOfDocx4j() {
        return this.r;
    }

    public Docx4jText() {
        this.r = FACTORY.createR();
    }

    public Docx4jText(R r, String text) {
        this.r = r;
        this.text = text;
    }

    /**
     * 获取或创建运行属性
     * @return 运行属性
     */
    private RPr getOrCreateRpr() {
        if (r.getRPr() == null) {
            RPr rPr = FACTORY.createRPr();
            r.setRPr(rPr);
            return rPr;
        }
        return r.getRPr();
    }
}
