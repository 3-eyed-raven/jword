package net.jsrbc.jword.core.api;

import net.jsrbc.jword.core.document.enums.TableJustification;
import net.jsrbc.jword.core.document.enums.TableWidthType;

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
     * @param styleId 段落整体样式
     * @param text 段落文字
     */
    JwordOperator addParagraph(String styleId, String text);

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
     */
    JwordOperator addCaptionLabel(String bookmarkName);

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

    JwordOperator addTableRow();
    JwordOperator addTableCell();
    JwordOperator addSection();

    /**
     * 保存文档
     * @param dest 保存路径
     * @param onProgress 进度事件
     * @param onError 错误事件
     * @param onComplete 完成事件
     */
    void saveAs(Path dest, Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete);
}
