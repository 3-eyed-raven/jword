package net.jsrbc.jword.core;

import net.jsrbc.jword.core.api.JwordOperator;
import net.jsrbc.jword.core.api.JwordLoader;
import net.jsrbc.jword.core.document.*;
import net.jsrbc.jword.core.document.enums.*;
import net.jsrbc.jword.core.document.visit.DocumentVisitor;
import net.jsrbc.jword.core.factory.AbstractJwordFactory;
import org.reactivestreams.Publisher;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;
import reactor.core.publisher.MonoSink;
import reactor.util.function.Tuple2;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.*;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Consumer;

/**
 * Jword Api
 * 非线程安全对象，单个实例请勿在多线程中操作
 * word模板文件，请保留至少一个分节内容，即sectPr
 * @author ZZZ on 2021/2/4 9:52
 * @version 1.0
 */
public class Jword implements JwordLoader, JwordOperator {

    /** Jword抽象工厂 */
    private final AbstractJwordFactory factory;

    /** ID生成器，文档内递增 */
    private final AtomicLong idGenerator = new AtomicLong(0);

    /** 文档命令集合 */
    private final List<Publisher<?>> commands = new ArrayList<>();

    /** 文档上下文 */
    private final Context context = new Context();

    /**
     * jword简单工厂
     * @param factory 文档对象抽象工厂
     * @return Jword对象
     */
    public static JwordLoader create(AbstractJwordFactory factory) {
        return new Jword(factory);
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator load(Path templatePath) {
        addCommand(sink -> {
            try {
                Document document = factory.loadDocument(templatePath);
                this.context.setDocument(document);
                sink.success(true);
            } catch (IOException e) {
                sink.error(e);
            }
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator walkBody(DocumentVisitor visitor) {
        addCommand(() -> this.context.getDocument().walkBody(visitor));
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator walkHeader(DocumentVisitor visitor) {
        addCommand(() -> this.context.getDocument().walkHeader(visitor));
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator walkFooter(DocumentVisitor visitor) {
        addCommand(() -> this.context.getDocument().walkFooter(visitor));
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addText(String text) {
        addCommand(() -> {
            Paragraph p = this.context.getCurrentParagraph();
            if (p == null) throw new IllegalStateException("paragraph is not exists");
            Text t = factory.createText();
            t.setValue(text);
            p.addText(t);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addStyledText(String styleId, String text) {
        addCommand(() -> {
            Paragraph p = this.context.getCurrentParagraph();
            if (p == null) throw new IllegalStateException("paragraph is not exists");
            Text t = factory.createText();
            t.setStyleId(styleId);
            t.setValue(text);
            p.addText(t);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addParagraph(String styleId) {
        addCommand(() -> {
            Paragraph paragraph = this.factory.createParagraph();
            paragraph.setStyleId(styleId);
            this.context.getDocument().addParagraph(paragraph);
            this.context.setCurrentParagraph(paragraph);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator createCaptionLabel(String bookmarkName, String label, int sequence) {
        return createCaptionLabel(bookmarkName, label, null, null, sequence);
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator createCaptionLabel(String bookmarkName, String label, String chapterStyleId, String chapterNumber, int sequence) {
        addCommand(() -> {
            CaptionLabel captionLabel = factory.createCaptionLabel();
            captionLabel.setBookmark(idGenerator.getAndIncrement(), bookmarkName);
            captionLabel.setLabel(label);
            captionLabel.setChapter(chapterStyleId, chapterNumber);
            captionLabel.setSequence(sequence);
            captionLabel.setSequence(sequence);
            this.context.putCaptionLabel(bookmarkName, captionLabel);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addReference(String bookmarkName) {
        addCommand(() -> {
            CaptionLabel captionLabel = this.context.getCaptionLabel(bookmarkName);
            if (captionLabel == null) throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Paragraph p = this.context.getCurrentParagraph();
            if (p == null) throw new IllegalStateException("paragraph is not exists");
            Reference reference = factory.createReference();
            reference.referTo(captionLabel);
            p.addReference(reference);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addCaptionLabel(String bookmarkName, String desc) {
        addCommand(() -> {
            CaptionLabel captionLabel = this.context.getCaptionLabel(bookmarkName);
            if (captionLabel == null) throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Paragraph p = this.context.getCurrentParagraph();
            if (p == null) throw new IllegalStateException("paragraph is not exists");
            p.addCaptionLabel(captionLabel);
            Text text = this.factory.createText();
            text.setValue("  " + desc);
            p.addText(text);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addDrawing(Path path, double width, double height) {
        addCommand(sink -> {
            Paragraph p = this.context.getCurrentParagraph();
            if (p == null)
                throw new IllegalStateException("paragraph is not exists");
            try {
                this.context.getDocument().addDrawing(p, path, width, height);
                sink.success(true);
            } catch (IOException e) {
                sink.error(e);
            }
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addTable(String styleId) {
        return addTable(styleId, 100, TableWidthType.PCT, TableJustification.CENTER);
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addTable(String styleId, double width, TableWidthType widthType, TableJustification justification) {
        addCommand(() -> {
            Table table = factory.createTable();
            table.setStyle(styleId);
            table.setWidth(width, widthType);
            table.setJustification(justification);
            this.context.getDocument().addTable(table);
            this.context.setCurrentTable(table);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addTableRow() {
        addCommand(() -> {
            Table table = this.context.getCurrentTable();
            if (table == null) throw new IllegalStateException("table is not exists");
            TableRow tableRow = factory.createTableRow();
            table.addRow(tableRow);
            this.context.setCurrentRow(tableRow);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addTableCell(double width, TableWidthType widthType, VerticalAlignType alignType, Integer span, VerticalMergeType mergeType) {
        addCommand(() -> {
            TableRow row = this.context.getCurrentRow();
            if (row == null) throw new IllegalStateException("table row is not exists");
            TableCell cell = this.factory.createTableCell();
            cell.setCellWidth(width, widthType);
            if (alignType != null) cell.setVerticalAlignType(alignType);
            if (span != null) cell.setGridSpan(span);
            if (mergeType != null) cell.setVerticalMergeType(mergeType);
            row.addCell(cell);
            this.context.setCurrentCell(cell);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addTextToCell(String styleId, String text) {
        addCommand(() -> {
            TableCell cell = this.context.getCurrentCell();
            if (cell == null) throw new IllegalStateException("table cell is not exists");
            Paragraph p = this.factory.createParagraph();
            p.setStyleId(styleId);
            Text t = this.factory.createText();
            t.setValue(text);
            p.addText(t);
            cell.addParagraph(p);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addCaptionLabelToCell(String styleId, String bookmarkName, String desc) {
        addCommand(() -> {
            TableCell cell = this.context.getCurrentCell();
            if (cell == null) throw new IllegalStateException("table cell is not exists");
            CaptionLabel label = this.context.getCaptionLabel(bookmarkName);
            if (label == null) throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Paragraph p = this.factory.createParagraph();
            p.setStyleId(styleId);
            p.addCaptionLabel(label);
            Text t = this.factory.createText();
            t.setValue("  " + desc);
            p.addText(t);
            cell.addParagraph(p);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addReferenceToCell(String styleId, String bookmarkName) {
        addCommand(() -> {
            TableCell cell = this.context.getCurrentCell();
            if (cell == null) throw new IllegalStateException("table cell is not exists");
            CaptionLabel label = this.context.getCaptionLabel(bookmarkName);
            if (label == null) throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Reference reference = this.factory.createReference();
            reference.referTo(label);
            Paragraph p = this.factory.createParagraph();
            p.setStyleId(styleId);
            p.addReference(reference);
            cell.addParagraph(p);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addDrawingToCell(String styleId, Path path, double width, double height) {
        addCommand(sink -> {
            TableCell cell = this.context.getCurrentCell();
            if (cell == null) throw new IllegalStateException("table cell is not exists");
            Paragraph p = this.factory.createParagraph();
            p.setStyleId(styleId);
            try {
                this.context.getDocument().addDrawing(p, path, width, height);
                cell.addParagraph(p);
                sink.success(true);
            } catch (IOException e) {
                sink.error(e);
            }
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addSection() {
        addCommand(() -> {
            Paragraph p = this.factory.createParagraph();
            Section section = this.factory.createSection();
            p.setSection(section);
            this.context.getDocument().addParagraph(p);
            this.context.setCurrentSection(section);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator setPageSize(PageSize pageSize) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.setPageSize(pageSize);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator setPageOrientation(PageOrientation orientation) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.setPageOrientation(orientation);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator setPageMargin(double top, double right, double bottom, double left) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.setPageMargin(top, right, bottom, left);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator setHeaderFootMargin(double headerMargin, double footerMargin) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.setHeaderMargin(headerMargin);
            section.setFooterMargin(footerMargin);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addHeaderReference(String id, HeaderFooterType headerFooterType) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.addHeaderReference(id, headerFooterType);
        });
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public JwordOperator addFooterReference(String id, HeaderFooterType headerFooterType) {
        addCommand(() -> {
            Section section = this.context.getCurrentSection();
            if (section == null)
                throw new IllegalStateException("section is not exists");
            section.addFooterReference(id, headerFooterType);
        });
        return this;
    }

    @Override
    public JwordOperator updateTableOfContent() {
        addCommand(() -> this.context.getDocument().updateTableOfContent());
        return this;
    }

    /** {@inheritDoc} */
    @Override
    public void saveAs(Path dest, Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete) {
        addCommand(sink -> {
            try {
                this.context.getDocument().saveAs(dest);
                sink.success(true);
            } catch (IOException e) {
                sink.error(e);
            }
        });
        finish(onProgress, onError, onComplete);
    }

    /** {@inheritDoc} */
    @Override
    public void saveAs(OutputStream os, Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete) {
        addCommand(sink -> {
            try {
                this.context.getDocument().saveAs(os);
                sink.success(true);
            } catch (IOException e) {
                sink.error(e);
            }
        });
        finish(onProgress, onError, onComplete);
    }

    private Jword(AbstractJwordFactory factory) {
        this.factory = factory;
    }

    /**
     * 添加命令
     * @param consumer sink消费
     */
    private <T> void addCommand(Consumer<MonoSink<T>> consumer) {
        this.commands.add(Mono.create(consumer));
    }

    /**
     * 添加命令
     * @param runnable 运行回调
     */
    private void addCommand(Runnable runnable) {
        this.commands.add(Mono.create(sink -> {
            runnable.run();
            sink.success(true); // 这里加上返回值，才能触发flux的OnNext事件
        }));
    }

    /**
     * 完成文档
     * @param onProgress 进度事件
     * @param onError 错误事件
     * @param onComplete 完成事件
     */
    private void finish(Consumer<? super Integer> onProgress, Consumer<? super Throwable> onError, Runnable onComplete) {
        Flux
                .concat(commands)
                .index()
                .map(Tuple2::getT1)
                .map(this::calcProgress)
                .distinctUntilChanged()
                .subscribe(onProgress, onError, onComplete);
    }

    /**
     * 计算进度
     * @param index 当前位置
     * @return 进度，总共100
     */
    private int calcProgress(long index) {
        if (commands.isEmpty()) return 100;
        return (int) (index + 1) * 100 / commands.size();
    }

    /**
     * JWord上下文
     */
    private static class Context {

        /** 当前文档 */
        private Document document;

        /** 最近一次操作的段落 */
        private Paragraph currentParagraph;

        /** 最近一次操作的章节 */
        private Section currentSection;

        /** 最近一次操作的表格 */
        private Table currentTable;

        /** 最近一次操作的表格行 */
        private TableRow currentRow;

        /** 最近一次操作的表格单元格 */
        private TableCell currentCell;

        private final Map<String, CaptionLabel> captionLabelMap = new HashMap<>();

        public Document getDocument() {
            return document;
        }

        public void setDocument(Document document) {
            this.document = document;
        }

        public Paragraph getCurrentParagraph() {
            return currentParagraph;
        }

        public void setCurrentParagraph(Paragraph currentParagraph) {
            this.currentParagraph = currentParagraph;
        }

        public Section getCurrentSection() {
            return currentSection;
        }

        public void setCurrentSection(Section currentSection) {
            this.currentSection = currentSection;
        }

        public void putCaptionLabel(String bookmarkName, CaptionLabel captionLabel) {
            captionLabelMap.put(bookmarkName, captionLabel);
        }

        public CaptionLabel getCaptionLabel(String bookmarkName) {
            return captionLabelMap.get(bookmarkName);
        }

        public Table getCurrentTable() {
            return currentTable;
        }

        public void setCurrentTable(Table currentTable) {
            this.currentTable = currentTable;
        }

        public TableRow getCurrentRow() {
            return currentRow;
        }

        public void setCurrentRow(TableRow currentRow) {
            this.currentRow = currentRow;
        }

        public TableCell getCurrentCell() {
            return currentCell;
        }

        public void setCurrentCell(TableCell currentCell) {
            this.currentCell = currentCell;
        }
    }
}
