package net.jsrbc.jword.core;

import net.jsrbc.jword.core.api.JwordOperator;
import net.jsrbc.jword.core.api.JwordLoader;
import net.jsrbc.jword.core.document.*;
import net.jsrbc.jword.core.document.enums.TableJustification;
import net.jsrbc.jword.core.document.enums.TableWidthType;
import net.jsrbc.jword.core.factory.AbstractJwordFactory;
import org.reactivestreams.Publisher;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;
import reactor.core.publisher.MonoSink;
import reactor.util.function.Tuple2;

import java.io.IOException;
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

    private final AtomicLong idGenerator = new AtomicLong(0);

    private final List<Publisher<?>> commands = new ArrayList<>();

    private final Context context = new Context();

    /**
     * jword简单工厂
     * @param factory 文档对象抽象工厂
     * @return Jword对象
     */
    public static JwordLoader create(AbstractJwordFactory factory) {
        return new Jword(factory);
    }

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

    @Override
    public JwordOperator addText(String text) {
        addCommand(() -> {
            Paragraph paragraph = getOrCreateCurrentParagraph();
            paragraph.addText(text);
        });
        return this;
    }

    @Override
    public JwordOperator addStyledText(String styleId, String text) {
        addCommand(() -> {
            Paragraph p = getOrCreateCurrentParagraph();
            p.addStyledText(styleId, text);
        });
        return this;
    }

    @Override
    public JwordOperator addParagraph(String styleId, String text) {
        addCommand(() -> {
            Paragraph paragraph = factory.createParagraph();
            paragraph.setStyleId(styleId);
            paragraph.addText(text);
            this.context.getDocument().addParagraph(paragraph);
            this.context.setCurrentParagraph(paragraph);
        });
        return this;
    }

    @Override
    public JwordOperator createCaptionLabel(String bookmarkName, String label, int sequence) {
        addCommand(() -> {
            CaptionLabel captionLabel = factory.createCaptionLabel();
            captionLabel.setBookmark(idGenerator.getAndIncrement(), bookmarkName);
            captionLabel.setLabel(label);
            captionLabel.setSequence(sequence);
            this.context.putCaptionLabel(bookmarkName, captionLabel);
        });
        return this;
    }

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

    @Override
    public JwordOperator addReference(String bookmarkName) {
        addCommand(() -> {
            CaptionLabel captionLabel = this.context.getCaptionLabel(bookmarkName);
            if (captionLabel == null)
                throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Paragraph p = getOrCreateCurrentParagraph();
            Reference reference = factory.createReference();
            reference.referTo(captionLabel);
            p.addReference(reference);
        });
        return this;
    }

    @Override
    public JwordOperator addCaptionLabel(String bookmarkName) {
        addCommand(() -> {
            CaptionLabel captionLabel = this.context.getCaptionLabel(bookmarkName);
            if (captionLabel == null)
                throw new NoSuchElementException(String.format("bookmarkName %s is not exists", bookmarkName));
            Paragraph p = getOrCreateCurrentParagraph();
            p.addCaptionLabel(captionLabel);
        });
        return this;
    }

    @Override
    public JwordOperator addTable(String styleId) {
        addCommand(() -> {
            Table table = factory.createTable();
            table.setStyle(styleId);
            table.setWidth(100, TableWidthType.PCT);
            table.setJustification(TableJustification.CENTER);
            this.context.getDocument().addTable(table);
            this.context.setCurrentTable(table);
        });
        return this;
    }

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

    @Override
    public JwordOperator addTableRow() {
        addCommand(() -> {
            Table table = this.context.getCurrentTable();
            if (table == null) throw new IllegalStateException("table is not exists");
            TableRow tableRow = factory.createTableRow();
            table.addRow(tableRow);
            this.context.setCurrentTable(table);
        });
        return this;
    }

    @Override
    public JwordOperator addTableCell() {
        return this;
    }

    @Override
    public JwordOperator addSection() {
        return this;
    }

    /**
     * 保存至指定路径
     * @param dest 目标路径
     */
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
     * 获取或者创建最近使用的段落
     * @return 段落
     */
    private Paragraph getOrCreateCurrentParagraph() {
        Paragraph p;
        if (this.context.getCurrentParagraph() != null) {
            p = this.context.getCurrentParagraph();
        } else {
            p = factory.createParagraph();
            this.context.getDocument().addParagraph(p);
            this.context.setCurrentParagraph(p);
        }
        return p;
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
     * 上下文
     */
    private static class Context {
        private Document document;
        private Paragraph currentParagraph;
        private Table currentTable;
        private TableRow currentRow;
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

        public Map<String, CaptionLabel> getCaptionLabelMap() {
            return captionLabelMap;
        }
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
}
