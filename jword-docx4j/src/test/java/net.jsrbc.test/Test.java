package net.jsrbc.test;

import net.jsrbc.jword.core.document.CaptionLabel;
import net.jsrbc.jword.core.document.Document;
import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.Reference;
import net.jsrbc.jword.docx4j.document.Docx4jCaptionLabel;
import net.jsrbc.jword.docx4j.document.Docx4jDocument;
import net.jsrbc.jword.docx4j.document.Docx4jParagraph;
import net.jsrbc.jword.docx4j.document.Docx4jReference;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.math.BigInteger;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * @author ZZZ on 2021/1/25 15:23
 * @version 1.0
 */
public class Test {

    private final static Path source = Paths.get("C:/Users/ZZZ/Desktop/template.docx");
    private final static Path dest = Paths.get("C:/Users/ZZZ/Desktop/test.docx");

    public static void main(String[] args) throws Throwable {
        testFrame();
    }

    private static void testFrame() throws Throwable {
        Document document = Docx4jDocument.load(source);

        Paragraph p = new Docx4jParagraph();
        p.setStyleId("1");
        p.addText("GY001 K1+94 GY田唐桥");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("2");
        p.addText("总体情况");
        document.addParagraph(p);

        CaptionLabel label1 = new Docx4jCaptionLabel(1, "bridgeTable");
        label1.setLabel("表");
        label1.setChapter("2", "1.1");
        label1.setSequence(1);
        CaptionLabel label2 = new Docx4jCaptionLabel(2, "frontImage");
        label2.setLabel("图");
        label2.setChapter("2", "1.1");
        label2.setSequence(1);
        CaptionLabel label3 = new Docx4jCaptionLabel(3, "sideImage");
        label3.setLabel("图");
        label3.setChapter("2", "1.1");
        label3.setSequence(2);

        p = new Docx4jParagraph();
        p.setStyleId("a");
        p.addText("GY田唐桥为单幅桥，相关信息如");
        Reference reference = new Docx4jReference();
        reference.referTo(label1);
        p.addReference(reference);
        p.addText("所示，桥梁正面照及侧面照如");
        reference = new Docx4jReference();
        reference.referTo(label2);
        p.addReference(reference);
        p.addText("与");
        reference = new Docx4jReference();
        reference.referTo(label3);
        p.addReference(reference);
        p.addText("所示。");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("a");
        p.addText("本次检查GY田唐桥评分为81分，等级为2类。");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label1);
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label2);
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label3);
        document.addParagraph(p);

        document.saveAs(dest);
    }

    private static void testRaw() throws Throwable {
        WordprocessingMLPackage wml = WordprocessingMLPackage.load(source.toFile());
        MainDocumentPart body = wml.getMainDocumentPart();
        body.addStyledParagraphOfText("1", "GY001 K1+94 GY田唐桥");
        // 添加书签
        ObjectFactory factory = Context.getWmlObjectFactory();
        P p = body.addParagraphOfText("");
        PPr pPr = factory.createPPr();
        PPrBase.PStyle pStyle = factory.createPPrBasePStyle();
        pStyle.setVal("afc");
        pPr.setPStyle(pStyle);
        p.setPPr(pPr);
        // 书签开始
        CTBookmark start = factory.createCTBookmark();
        start.setId(BigInteger.ZERO);
        start.setName("_Ref62722846");
        JAXBElement<CTBookmark> s = factory.createBodyBookmarkStart(start);
        //标签
        R title = factory.createR();
        Text text = factory.createText();
        text.setValue("表");
        title.getContent().add(text);
        //编号
        R frstart = factory.createR();
        FldChar fstart = factory.createFldChar();
        fstart.setFldCharType(STFldCharType.BEGIN);
        frstart.getContent().add(fstart);

        R number = factory.createR();
        text = factory.createText();
        text.setValue("STYLEREF 1 \\s");
        JAXBElement<Text> numberText = factory.createRInstrText(text);
        number.getContent().add(numberText);

        R sep1 = factory.createR();
        FldChar sep1C = factory.createFldChar();
        sep1C.setFldCharType(STFldCharType.SEPARATE);
        sep1.getContent().add(sep1C);

        R n = factory.createR();
        text = factory.createText();
        text.setValue("1");
        n.getContent().add(text);

        R frend = factory.createR();
        FldChar fend = factory.createFldChar();
        fend.setFldCharType(STFldCharType.END);
        frend.getContent().add(fend);

        // 书签结束
        CTMarkupRange end = factory.createCTBookmark();
        end.setId(BigInteger.ZERO);
        JAXBElement<CTMarkupRange> e = factory.createBodyBookmarkEnd(end);

        //书签内容加入段落
        p.getContent().add(s);
        p.getContent().add(title);
        p.getContent().add(frstart);
        p.getContent().add(number);
        p.getContent().add(sep1);
        p.getContent().add(n);
        p.getContent().add(frend);
        p.getContent().add(e);

        wml.save(dest.toFile());
    }
}
