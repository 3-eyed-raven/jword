package net.jsrbc.jword.docx4j.factory;

import net.jsrbc.jword.core.document.*;
import net.jsrbc.jword.core.factory.AbstractJwordFactory;
import net.jsrbc.jword.docx4j.document.*;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;

/**
 * Docx4j Jword工厂
 * @author ZZZ on 2021/2/4 11:21
 * @version 1.0
 */
public class Docx4jJwordFactory implements AbstractJwordFactory {

    /** {@inheritDoc} */
    @Override
    public Document loadDocument(Path path) throws IOException {
        return Docx4jDocument.load(path);
    }

    /** {@inheritDoc} */
    @Override
    public Document loadDocument(InputStream in) throws IOException {
        return Docx4jDocument.load(in);
    }

    /** {@inheritDoc} */
    @Override
    public Paragraph createParagraph() {
        return new Docx4jParagraph();
    }

    /** {@inheritDoc} */
    @Override
    public Text createText() {
        return new Docx4jText();
    }

    /** {@inheritDoc} */
    @Override
    public CaptionLabel createCaptionLabel() {
        return new Docx4jCaptionLabel();
    }

    /** {@inheritDoc} */
    @Override
    public Reference createReference() {
        return new Docx4jReference();
    }

    /** {@inheritDoc} */
    @Override
    public Section createSection() {
        return new Docx4jSection();
    }

    /** {@inheritDoc} */
    @Override
    public Table createTable() {
        return new Docx4jTable();
    }

    /** {@inheritDoc} */
    @Override
    public TableRow createTableRow() {
        return new Docx4jTableRow();
    }

    /** {@inheritDoc} */
    @Override
    public TableCell createTableCell() {
        return new Docx4jTableCell();
    }
}
