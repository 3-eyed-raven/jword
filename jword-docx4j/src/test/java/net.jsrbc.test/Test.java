package net.jsrbc.test;

import net.jsrbc.jword.core.Jword;
import net.jsrbc.jword.core.document.Text;
import net.jsrbc.jword.core.document.enums.*;
import net.jsrbc.jword.core.document.visit.DocumentVisitResult;
import net.jsrbc.jword.core.document.visit.DocumentVisitor;
import net.jsrbc.jword.docx4j.factory.Docx4jJwordFactory;

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
//        test();
        testRaw();
    }

    private static void test() throws Exception {
        Jword.create(new Docx4jJwordFactory())
                .load(source)
                .walkHeader(new DocumentVisitor() {
                    @Override
                    public DocumentVisitResult visitText(Text text) {
                        String str = text.getValue();
                        if (str.contains("报告编号")) {
                            text.setValue("测试桥梁");
                        }
                        return DocumentVisitResult.CONTINUE;
                    }
                })
                .addParagraph("1").addText("GY001 K1+94 GY田唐桥")
                .addParagraph("2").addText("总体情况")
                .createCaptionLabel("bridgeTable", "表", "2", "1.1", 1)
                .createCaptionLabel("frontImage", "图", "2", "1.1", 1)
                .createCaptionLabel("sideImage" , "图", "2", "1.1", 2)
                .addParagraph("a").addText("GY田唐桥为单幅桥，相关信息如")
                .addReference("bridgeTable")
                .addText("所示，桥梁正面照及侧面照如")
                .addReference("frontImage").addText("与")
                .addReference("sideImage").addText("所示")
                .addParagraph("a").addText("本次检查GY田唐桥评分为81分，等级为2类。")
                .addParagraph("afc")
                .addCaptionLabel("bridgeTable", "桥梁基本信息一览表")
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
                .addParagraph("a")
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
                .addParagraph("2").addText("检查评定结果")
                .addParagraph("3").addText("桥梁技术状况评定结果")
                .createCaptionLabel("evaluationTable", "表", "3", "1.2.1", 1)
                .addParagraph("a").addText("根据现场检查结果，依据《公路桥梁技术状况评定标准》（JTG/T H21-2011），对桥梁各评定单元进行技术状况评定，结果列于")
                .addReference("evaluationTable").addText("。")
                .addParagraph("afc")
                .addCaptionLabel("evaluationTable", "GY田唐桥技术状况评定结果表")
                .addTable("ac")
                .addTableRow()
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "评定单元名称")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "对应孔跨")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "部位")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "评分")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "部位技术状况等级")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "综合评分")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "评定方法类型")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ad", "单元技术状况等级")
                .addTableRow()
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.RESTART)
                .addTextToCell("ae", "单幅(整体现浇板梁)")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.RESTART)
                .addTextToCell("ae", "1")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "上部结构")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "79.0")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "3类")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.RESTART)
                .addTextToCell("ae", "80.7")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.RESTART)
                .addTextToCell("ae", "综合评定")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.RESTART)
                .addTextToCell("ae", "2类")
                .addTableRow()
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "下部结构")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "95.9")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "1类")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableRow()
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "桥面系")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "53.8")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, null)
                .addTextToCell("ae", "4类")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .addTableCell(12.5, TableWidthType.PCT, VerticalAlignType.CENTER, null, VerticalMergeType.CONTINUE)
                .addTextToCell("ae", "")
                .updateTableOfContent()
                .saveAs(dest, System.out::println, Throwable::printStackTrace, () -> System.out.println("完成"));
    }

    public static void testRaw() throws Throwable {
        Jword.create(new Docx4jJwordFactory())
                .load(source)
                .walkHeader(new DocumentVisitor() {
                    @Override
                    public DocumentVisitResult visitText(Text text) {
                        String str = text.getValue();
                        System.out.println(str);
                        return DocumentVisitResult.CONTINUE;
                    }
                })
                .read(System.out::println, Throwable::printStackTrace, () -> System.out.println("完成"));
    }
}
