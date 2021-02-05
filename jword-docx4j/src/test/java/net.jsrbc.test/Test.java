package net.jsrbc.test;

import net.jsrbc.jword.core.Jword;
import net.jsrbc.jword.core.document.enums.*;
import net.jsrbc.jword.docx4j.factory.Docx4jJwordFactory;
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
    private final static Path frontImage = Paths.get("C:/Users/ZZZ/Desktop/1.jpg");
    private final static Path sideImage = Paths.get("C:/Users/ZZZ/Desktop/2.jpg");

    public static void main(String[] args) throws Throwable {
        test();
    }

    private static void test() throws Exception {
        Jword.create(new Docx4jJwordFactory())
                .load(source)
                .addParagraph("1", "GY001 K1+94 GY田唐桥")
                .addParagraph("2", "总体情况")
                .createCaptionLabel("bridgeTable", "表", "2", "1.1", 1)
                .createCaptionLabel("frontImage", "图", "2", "1.1", 1)
                .createCaptionLabel("sideImage" , "图", "2", "1.1", 2)
                .addParagraph("a", "GY田唐桥为单幅桥，相关信息如")
                .addReference("bridgeTable")
                .addText("所示，桥梁正面照及侧面照如")
                .addReference("frontImage").addText("与")
                .addReference("sideImage").addText("所示")
                .addParagraph("a", "本次检查GY田唐桥评分为81分，等级为2类。")
                .addParagraph("afc", "")
                .addCaptionLabel("bridgeTable").addText("  桥梁基本信息一览表")
                .addTable("ac")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "路线名称")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "仲田线")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥梁名称")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "GY田唐桥")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "中心桩号")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "K1+94")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥梁编码")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "建成时间")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "上部结构材料")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "钢筋混凝土")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "上部结构类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "整体现浇板梁")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥墩类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥台类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "重力式桥台、重力式桥台")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "支座类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "伸缩缝类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥面铺装类型")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "水泥混凝土")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥面总宽(m)")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "7.0")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "上次检查日期")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "/")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "本次检查日期")
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "2020-09-12")
                .addTableRow()
                .addTableCell(16.66, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "桥跨组成")
                .addTextToCell("ad", "(孔*m)")
                .addTableCell(100 - 16.66, TableWidthType.PCT, VerticalAlignType.CENTER, 5, null)
                .addTextToCell("ae", "1*5")
                .addParagraph("a", "")
                .addTable("a1")
                .addTableRow()
                .addTableCell(50, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addDrawingToCell("ad", frontImage, 7.5, 5)
                .addTableCell(50, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addDrawingToCell("ad", sideImage, 7.5, 5)
                .addTableRow()
                .createCaptionLabel("frontImage", "图", "2", "1.1", 1)
                .createCaptionLabel("sideImage", "图", "2", "1.1", 2)
                .addTableCell(50, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addCaptionLabelToCell("ad", "frontImage", "桥梁正面照")
                .addTableCell(50, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addCaptionLabelToCell("ad", "sideImage", "桥梁侧面照")
                .addParagraph("a", "")
                .addSection()
                .setPageSize(PageSize.A3)
                .setPageOrientation(PageOrientation.LANDSCAPE)
                .setPageMargin(2.5, 2.5, 2.5, 2.5)
                .setHeaderFootMargin(1.5, 1.5)
                .addParagraph("a", "")
                .addDrawing(frontImage, 10, 10)
                .saveAs(dest, System.out::println, Throwable::printStackTrace, () -> System.out.println("完成"));
    }
}
