package net.jsrbc.jword.core.api;

import net.jsrbc.jword.core.document.enums.*;
import net.jsrbc.jword.core.document.visit.DocumentVisitor;
import reactor.util.annotation.Nullable;

import java.io.OutputStream;
import java.nio.file.Path;
import java.util.function.Consumer;

/**
 * Jword 操作
 * 非线程安全对象，单个实例请勿在多线程中操作
 * word模板文件，请保留至少一个分节内容，即sectPr
 * @author ZZZ on 2021/2/4 14:37
 * @version 1.0
 */
public interface JwordOperator {

    /**
     * 遍历文档文档内容
     * @param visitor 访问者
     */
    JwordOperator walkBody(DocumentVisitor visitor);

    /**
     * 遍历文档文档页眉
     * @param visitor 访问者
     */
    JwordOperator walkHeader(DocumentVisitor visitor);

    /**
     * 遍历文档文档页脚
     * @param visitor 访问者
     */
    JwordOperator walkFooter(DocumentVisitor visitor);

    /**
     * 在最近的生成的段落上添加文字
     * 如果段落不存在，则会创建一个段落，并把文字添加上去
     * @param text 添加到段落的文字
     */
    JwordOperator addText(String text);

    /**
     * 在最近的生成的段落上添加带样式文字
     * 如果段落不存在，则会创建一个段落，并把文字添加上去
     * @param styleId 文字样式
     * @param text 添加到段落的文字
     */
    JwordOperator addStyledText(String styleId, String text);

    /**
     * 添加一个段落
     * @param styleId 段落样式
     */
    JwordOperator addParagraph(String styleId);

    /**
     * 创建一个不带章节编号的题注，注意此时并未添加到文档中，需要调用addCaptionLabel进行添加
     * @param bookmarkName 书签名，文档内应该唯一
     * @param label 题注标签
     * @param sequence 序号，如：1
     */
    JwordOperator createCaptionLabel(String bookmarkName, String label, int sequence);

    /**
     * 创建一个带章节编号的题注，注意此时并未添加到文档中，需要调用addCaptionLabel进行添加
     * @param bookmarkName 书签名，文档内应该唯一
     * @param label 题注标签
     * @param chapterStyleId 章节起始样式ID
     * @param chapterNumber 章节编号，如：1.1.1
     * @param sequence 序号，如：1
     */
    JwordOperator createCaptionLabel(String bookmarkName, String label, String chapterStyleId, String chapterNumber, int sequence);

    /**
     * 在最近的段落上添加交叉引用，前提条件是该题注已经创建
     * 如果段落不存在，则会创建一个段落，并把交叉引用添加上去
     * @param bookmarkName 书签名
     */
    JwordOperator addReference(String bookmarkName);

    /**
     * 在最近的段落上添加题注，前提条件是该题注已经创建
     * 如果段落不存在，则会创建一个段落，并把题注添加上去
     * @param bookmarkName 书签名
     * @param desc 书签描述
     */
    JwordOperator addCaptionLabel(String bookmarkName, String desc);

    /**
     * 添加图画到最近一次生成的段落
     * @param path 图画路径
     * @param width 图画宽度，单位：cm
     * @param height 图画高度，单位：cm
     */
    JwordOperator addDrawing(Path path, double width, double height);

    /**
     * 添加表格，表格宽度与文档宽度一致
     * @param styleId 表格样式ID
     */
    JwordOperator addTable(String styleId);

    /**
     * 添加表格
     * @param styleId 表格样式ID
     * @param width 表格宽度
     * @param widthType 宽度类型
     * @param justification 表格横向对齐方式
     */
    JwordOperator addTable(String styleId, double width, TableWidthType widthType, TableJustification justification);

    /**
     * 添加表格行到最近创建的表格
     */
    JwordOperator addTableRow();

    /**
     * 添加文字单元格至最近一次创建的表格行
     * @param width 单元格宽度
     * @param widthType 宽度类型
     * @param alignType 垂直对齐方式
     * @param span 横向合并网格数
     * @param mergeType 单元格行级合并，restart为合并起点
     */
    JwordOperator addTableCell(double width, TableWidthType widthType, @Nullable VerticalAlignType alignType, @Nullable Integer span, @Nullable VerticalMergeType mergeType);

    /**
     * 添加文字到最近一次添加的单元格
     * @param styleId 段落样式ID
     * @param text 文字格式
     */
    JwordOperator addTextToCell(String styleId, String text);

    /**
     * 添加题注到最近一次生成的单元格
     * @param styleId 段落样式ID
     * @param bookmarkName 书签名
     * @param desc 描述性文字
     */
    JwordOperator addCaptionLabelToCell(String styleId, String bookmarkName, String desc);

    /**
     * 添加交叉引用到最近一次生成的单元格
     * @param styleId 段落样式ID
     * @param bookmarkName 书签名
     */
    JwordOperator addReferenceToCell(String styleId, String bookmarkName);

    /**
     * 添加图画到最近一次生成的单元格
     * @param styleId 段落样式ID
     * @param path 图画路径
     * @param width 图画宽度，单位：cm
     * @param height 图画高度，单位：cm
     */
    JwordOperator addDrawingToCell(String styleId, Path path, double width, double height);

    /**
     * 添加分节符划分章节
     * 上一个分节符至设置本分节符的段落均属于一个章节
     */
    JwordOperator addSection();

    /**
     * 设置当前节页面大小
     * @param pageSize 页面大小
     */
    JwordOperator setPageSize(PageSize pageSize);

    /**
     * 设置当前节页面方向
     * @param orientation 方向
     */
    JwordOperator setPageOrientation(PageOrientation orientation);

    /**
     * 设置当前节页面边距
     * @param top 上边距，单位：cm
     * @param right 右边距，单位：cm
     * @param bottom 下边距，单位：cm
     * @param left 左边距，单位：cm
     */
    JwordOperator setPageMargin(double top, double right, double bottom, double left);

    /**
     * 设置当前节页眉页脚的边距
     * @param headerMargin 页眉边距，单位：cm
     * @param footerMargin 页脚边距，单位：cm
     */
    JwordOperator setHeaderFootMargin(double headerMargin, double footerMargin);

    /**
     * 设置当前节页眉引用
     * @param id 引用的页眉ID
     * @param headerFooterType 页眉类型
     */
    JwordOperator addHeaderReference(String id, HeaderFooterType headerFooterType);

    /**
     * 设置页脚引用
     * @param id 引用的页脚ID
     * @param headerFooterType 页脚类型
     */
    JwordOperator addFooterReference(String id, HeaderFooterType headerFooterType);

    /**
     * 更新目录
     */
    JwordOperator updateTableOfContent();

    /**
     * 仅读取文档
     * @param onProgress 进度事件
     * @param onError 错误事件
     * @param onComplete 完成事件
     */
    void read(Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete);

    /**
     * 另存文档
     * @param dest 保存路径
     * @param onProgress 进度事件
     * @param onError 错误事件
     * @param onComplete 完成事件
     */
    void saveAs(Path dest, Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete);

    /**
     * 另存文档
     * @param os 输出流
     * @param onProgress 进度事件
     * @param onError 错误事件
     * @param onComplete 完成事件
     */
    void saveAs(OutputStream os, Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete);
}
