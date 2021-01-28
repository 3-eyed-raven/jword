package net.jsrbc.jword.docx4j.factory;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;

/**
 * Docx4j工具
 * @author ZZZ on 2021/1/28 17:22
 * @version 1.0
 */
public final class Docx4jFactory {

    private final static ObjectFactory factory = Context.getWmlObjectFactory();

    public static R fldBegin() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.BEGIN);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    public static R fldSeparate() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.SEPARATE);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    public static R fldEnd() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.END);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    public static R text(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(text);
        r.getContent().add(t);
        return r;
    }

    public static R styledText(String styleId, String text) {
        R r = factory.createR();
        // 设置样式
        RPr rPr = factory.createRPr();
        RStyle style = factory.createRStyle();
        style.setVal(styleId);
        rPr.setRStyle(style);
        // 设置文字
        Text txt = factory.createText();
        txt.setValue(text);
        // 拼装
        r.setRPr(rPr);
        r.getContent().add(txt);
        return r;
    }

    public static R instrText(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(text);
        JAXBElement<Text> textElement = factory.createRInstrText(t);
        r.getContent().add(textElement);
        return r;
    }

    public static R preserveInstrText(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setSpace("preserve");
        t.setValue(text);
        JAXBElement<Text> textElement = factory.createRInstrText(t);
        r.getContent().add(textElement);
        return r;
    }

    public static R noBreakHyphen() {
        R r = factory.createR();
        R.NoBreakHyphen noBreakHyphen = factory.createRNoBreakHyphen();
        r.getContent().add(noBreakHyphen);
        return r;
    }

    private Docx4jFactory() {}
}
