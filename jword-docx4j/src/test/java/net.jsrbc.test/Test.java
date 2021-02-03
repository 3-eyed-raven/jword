package net.jsrbc.test;

import net.jsrbc.jword.core.document.*;
import net.jsrbc.jword.core.document.Document;
import net.jsrbc.jword.core.document.enums.*;
import net.jsrbc.jword.docx4j.document.*;
import net.jsrbc.jword.docx4j.util.UnitConverter;
import org.docx4j.dml.CTNonVisualDrawingProps;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.wordprocessingDrawing.CTEffectExtent;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.docx4j.wml.Drawing;

import java.nio.file.Path;
import java.nio.file.Paths;

import static net.jsrbc.jword.docx4j.util.UnitConverter.*;

/**
 * @author ZZZ on 2021/1/25 15:23
 * @version 1.0
 */
public class Test {

    private final static Path source = Paths.get("C:/Users/ZZZ/Desktop/template.docx");
    private final static Path dest = Paths.get("C:/Users/ZZZ/Desktop/test.docx");
    private final static Path imagePath = Paths.get("C:/Users/ZZZ/Desktop/1.jpg");
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
        firstSection.setPageMargin(2.5, 2.5, 2.5, 2.5);
        firstSection.setHeaderMargin(1.5);
        firstSection.setFooterMargin(1.5);

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
        p.addText("  桥梁基本信息一览表");
        document.addParagraph(p);

        document.addTable(createTable());

        p = new Docx4jParagraph();
        p.setSection(firstSection);
        document.addParagraph(p);

        document.addTable(createImageTable(label2, label3));

        document.saveAs(dest);
    }

    private static Table createTable() {
        Table table = new Docx4jTable();
        table.setStyle("ac");
        table.setWidth(100, TableWidthType.PCT);

        TableRow tableRow = new Docx4jTableRow();
        TableCell cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        Paragraph p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("路线名称");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("仲田线");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥梁名称");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("GY田唐桥");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("中心桩号");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("K1+94");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        //2=====================
        tableRow = new Docx4jTableRow();
        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥梁编码");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("建成时间");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("上部结构材料");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("钢筋混凝土");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        //3==================================
        tableRow = new Docx4jTableRow();
        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("上部结构类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("整体现浇板梁");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥墩类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥台类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("重力式桥台、重力式桥台");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        //4=================================================================
        tableRow = new Docx4jTableRow();
        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("支座类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("伸缩缝类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥面铺装类型");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("水泥混凝土");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        //5===========================================
        tableRow = new Docx4jTableRow();
        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥面总宽(m)");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("7");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("上次检查日期");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("/");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("本次检查日期");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("2020-09-12");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        // last
        tableRow = new Docx4jTableRow();
        cell = new Docx4jTableCell();
        cell.setCellWidth(16.6, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("桥梁孔跨组成");
        cell.addParagraph(p);
        p = new Docx4jParagraph();
        p.setStyleId("ad");
        p.addText("（孔*m）");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(83.4, TableWidthType.PCT);
        cell.setGridSpan(5);
        cell.setVerticalAlignType(VerticalAlignType.CENTER);
        p = new Docx4jParagraph();
        p.setStyleId("ae");
        p.addText("1*5");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        return table;
    }

    private static Table createImageTable(CaptionLabel label1, CaptionLabel label2) {
        Table table = new Docx4jTable();
        table.setStyle("a3");
        table.setWidth(100, TableWidthType.PCT);

        TableRow tableRow = new Docx4jTableRow();

        TableCell cell = new Docx4jTableCell();
        cell.setCellWidth(50, TableWidthType.PCT);
        cell.addParagraph(new Docx4jParagraph());
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(50, TableWidthType.PCT);
        cell.addParagraph(new Docx4jParagraph());
        tableRow.addCell(cell);

        table.addRow(tableRow);

        tableRow = new Docx4jTableRow();

        cell = new Docx4jTableCell();
        cell.setCellWidth(50, TableWidthType.PCT);
        Paragraph p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label1);
        p.addText("  桥梁正面照");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        cell = new Docx4jTableCell();
        cell.setCellWidth(50, TableWidthType.PCT);
        p = new Docx4jParagraph();
        p.setStyleId("afc");
        p.addCaptionLabel(label2);
        p.addText("  桥梁侧面照");
        cell.addParagraph(p);
        tableRow.addCell(cell);

        table.addRow(tableRow);
        return table;
    }

    private static void testRaw() throws Throwable {
        WordprocessingMLPackage wml = WordprocessingMLPackage.load(source.toFile());
        MainDocumentPart mainPart = wml.getMainDocumentPart();

        BinaryPartAbstractImage part = BinaryPartAbstractImage.createImagePart(wml, imagePath.toFile());
        Inline inline = part.createImageInline("a", "b", 1, 2, false);

        CTPositiveSize2D size = new CTPositiveSize2D();
        size.setCx(cmToEMUs(5));
        size.setCy(cmToEMUs(5));
        inline.setExtent(size);

        CTNonVisualDrawingProps docPr = new CTNonVisualDrawingProps();
        docPr.setId(1);

        P p = FACTORY.createP();
        R r = FACTORY.createR();
        p.getContent().add(r);
        Drawing drawing = FACTORY.createDrawing();
        r.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        mainPart.getContent().add(p);

        wml.save(dest.toFile());
    }
}
