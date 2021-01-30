package net.jsrbc.jword.docx4j.factory;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;

/**
 * Docx4j简单工厂
 * @author ZZZ on 2021/1/28 17:22
 * @version 1.0
 */
public final class Docx4jFactory {

    private final static ObjectFactory factory = Context.getWmlObjectFactory();

    /**
     * 创建域起始标签
     * @return 域起始标签
     */
    public static R createFldBegin() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.BEGIN);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    /**
     * 创建域分隔符
     * @return 创建域分隔符
     */
    public static R createFldSeparate() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.SEPARATE);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    /**
     * 创建域结束符
     * @return 域结束符
     */
    public static R createFldEnd() {
        FldChar fldChar = factory.createFldChar();
        fldChar.setFldCharType(STFldCharType.END);
        R r = factory.createR();
        r.getContent().add(fldChar);
        return r;
    }

    /**
     * 创建文本
     * @param text 文本内容
     * @return 文本
     */
    public static R createText(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(text);
        t.setSpace("preserve");
        r.getContent().add(t);
        return r;
    }

    /**
     * 创建带样式的文本
     * @param styleId 样式ID
     * @param text 文本内容
     * @return 样式文本
     */
    public static R createStyledText(String styleId, String text) {
        R r = factory.createR();
        // 设置样式
        RPr rPr = factory.createRPr();
        RStyle style = factory.createRStyle();
        style.setVal(styleId);
        rPr.setRStyle(style);
        // 设置文字
        Text txt = factory.createText();
        txt.setValue(text);
        txt.setSpace("preserve");
        // 拼装
        r.setRPr(rPr);
        r.getContent().add(txt);
        return r;
    }

    /**
     * 创建命令文本
     * @param text 文本内容
     * @return 命令文本
     */
    public static R createInstrText(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setValue(text);
        JAXBElement<Text> textElement = factory.createRInstrText(t);
        r.getContent().add(textElement);
        return r;
    }

    /**
     * 创建保留空白文字的命令文本
     * @param text 文本内容
     * @return 保留空白字段的文本
     */
    public static R createPreserveInstrText(String text) {
        R r = factory.createR();
        Text t = factory.createText();
        t.setSpace("preserve");
        t.setValue(text);
        JAXBElement<Text> textElement = factory.createRInstrText(t);
        r.getContent().add(textElement);
        return r;
    }

    /**
     * 创建超链接连接符
     * @return 超链接连接符
     */
    public static R createNoBreakHyphen() {
        R r = factory.createR();
        R.NoBreakHyphen noBreakHyphen = factory.createRNoBreakHyphen();
        r.getContent().add(noBreakHyphen);
        return r;
    }

    private Docx4jFactory() {}
}
