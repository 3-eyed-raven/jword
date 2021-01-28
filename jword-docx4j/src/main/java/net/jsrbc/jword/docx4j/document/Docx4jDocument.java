package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Document;
import net.jsrbc.jword.core.document.Paragraph;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;

/**
 * DOCX4J文档
 * @author ZZZ on 2021/1/27 14:33
 * @version 1.0
 */
public class Docx4jDocument implements Document {

    /** 文档来源 */
    private final Path source;

    /** 文档包 */
    private final WordprocessingMLPackage wml;

    /** 文档主体 */
    private final MainDocumentPart body;

    /**
     * 加载word文档
     * @param path 文件路径
     * @return 文档对象
     * @throws IOException 文件未找到或不可用时抛出异常
     */
    public static Docx4jDocument load(Path path) throws IOException {
        try {
            WordprocessingMLPackage wml = WordprocessingMLPackage.load(path.toFile());
            return new Docx4jDocument(path, wml);
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    /** {@inheritDoc} */
    @Override
    public void addParagraph(Paragraph paragraph) {
        if (!(paragraph instanceof Docx4jParagraph))
            throw new IllegalArgumentException("paragraph is not Docx4jParagraph");
        body.getContent().add(((Docx4jParagraph)paragraph).getParagraphOfDocx4j());
    }

    /** {@inheritDoc} */
    @Override
    public void save() throws IOException {
        try {
            this.wml.save(source.toFile());
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    /** {@inheritDoc} */
    @Override
    public void saveAs(Path dest) throws IOException {
        try {
            this.wml.save(dest.toFile());
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    /** {@inheritDoc} */
    @Override
    public void saveAs(OutputStream dest) throws IOException {
        try {
            this.wml.save(dest);
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    private Docx4jDocument(Path source, WordprocessingMLPackage wml) {
        this.source = source;
        this.wml = wml;
        this.body = wml.getMainDocumentPart();
    }
}
