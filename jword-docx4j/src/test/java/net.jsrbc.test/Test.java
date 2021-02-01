package net.jsrbc.test;

import net.jsrbc.jword.core.document.*;
import net.jsrbc.jword.core.document.Document;
import net.jsrbc.jword.core.document.enums.HeaderFooterType;
import net.jsrbc.jword.core.document.enums.PageSize;
import net.jsrbc.jword.core.document.enums.SectionType;
import net.jsrbc.jword.docx4j.document.*;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * @author ZZZ on 2021/1/25 15:23
 * @version 1.0
 */
public class Test {

    private final static Path source = Paths.get("C:/Users/ZZZ/Desktop/template.docx");
    private final static Path dest = Paths.get("C:/Users/ZZZ/Desktop/test.docx");
    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    public static void main(String[] args) throws Throwable {
        testFrame();
//        testRaw();
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

        Section firstSection = new Docx4jSection();
        firstSection.addHeaderReference("rId9", HeaderFooterType.DEFAULT);
        firstSection.addFooterReference("rId11", HeaderFooterType.DEFAULT);
        firstSection.setPageSize(PageSize.A4);
        firstSection.setPageMargin(1418, 1418, 1418, 1418);
        firstSection.setHeaderMargin(851);
        firstSection.setFooterMargin(851);

        Section defaultSection = new Docx4jSection();
        defaultSection.setPageSize(PageSize.A4);
        defaultSection.setType(SectionType.CONTINUOUS);
        defaultSection.setPageMargin(1418, 1418, 1418, 1418);
        defaultSection.setHeaderMargin(851);
        defaultSection.setFooterMargin(851);

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
        p.setSection(firstSection);
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("a");
        p.addText("本次检查GY田唐桥评分为81分，等级为2类。");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setSection(defaultSection);
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label1);
        p.addText("  桥梁基本信息一览表");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label2);
        p.addText("  桥梁正面照");
        document.addParagraph(p);

        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label3);
        p.addText("  桥梁侧面照");
        document.addParagraph(p);

        document.addSection(defaultSection);

        document.saveAs(dest);
    }

    private static void testRaw() throws Throwable {
    }
}
