package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Document;
import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.Section;
import net.jsrbc.jword.core.document.Table;
import net.jsrbc.jword.core.document.visit.DocumentVisitResult;
import net.jsrbc.jword.core.document.visit.DocumentVisitor;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.toc.TocException;
import org.docx4j.toc.TocGenerator;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicInteger;

import static net.jsrbc.jword.docx4j.util.UnitConverter.*;

/**
 * DOCX4J文档
 * @author ZZZ on 2021/1/27 14:33
 * @version 1.0
 */
public class Docx4jDocument implements Document {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    /** ID计数器 */
    private final AtomicInteger idGenerator = new AtomicInteger(0);

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
            return new Docx4jDocument(wml);
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    /**
     * 从流加载word文档
     * @param in 输入流
     * @return 文档对象
     * @throws IOException 文件未找到或不可用时抛出异常
     */
    public static Docx4jDocument load(InputStream in) throws IOException {
        try {
            WordprocessingMLPackage wml = WordprocessingMLPackage.load(in);
            return new Docx4jDocument(wml);
        } catch (Docx4JException e) {
            throw new IOException(e);
        }
    }

    /** {@inheritDoc} */
    @Override
    public void walkBody(DocumentVisitor visitor) {
        walk(this.body, visitor);
    }

    /** {@inheritDoc} */
    @Override
    public void walkHeader(DocumentVisitor visitor) {
        for (Relationship relationship : this.body.getRelationshipsPart().getRelationshipsByType(Namespaces.HEADER)) {
            HeaderPart headerPart = (HeaderPart) this.body.getRelationshipsPart().getPart(relationship);
            DocumentVisitResult result = walk(headerPart, visitor);
            if (result == DocumentVisitResult.TERMINATE) return;
        }
    }

    /** {@inheritDoc} */
    @Override
    public void walkFooter(DocumentVisitor visitor) {
        for (Relationship relationship : this.body.getRelationshipsPart().getRelationshipsByType(Namespaces.FOOTER)) {
            FooterPart footerPart = (FooterPart) this.body.getRelationshipsPart().getPart(relationship);
            DocumentVisitResult result = walk(footerPart, visitor);
            if (result == DocumentVisitResult.TERMINATE) return;
        }
    }

    /** {@inheritDoc} */
    @Override
    public void addTable(Table table) {
        if (!(table instanceof Docx4jTable))
            throw new IllegalArgumentException("table is not Docx4jTable");
        this.body.getContent().add(((Docx4jTable) table).getTableOfDocx4j());
    }

    /** {@inheritDoc} */
    @Override
    public void addParagraph(Paragraph paragraph) {
        if (!(paragraph instanceof Docx4jParagraph))
            throw new IllegalArgumentException("paragraph is not Docx4jParagraph");
        this.body.getContent().add(((Docx4jParagraph)paragraph).getParagraphOfDocx4j());
    }

    /** {@inheritDoc} */
    @Override
    public void addDrawing(Paragraph target, Path path, double width, double height) throws IOException {
        if (!(target instanceof Docx4jParagraph))
            throw new IllegalArgumentException("paragraph is not Docx4jParagraph");
        P p = ((Docx4jParagraph) target).getParagraphOfDocx4j();
        try {
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(this.wml, path.toFile());
            Inline inline = imagePart.createImageInline("picture", "error", idGenerator.getAndIncrement(), idGenerator.getAndIncrement(), false);
            // 设置图片大小
            CTPositiveSize2D size = new CTPositiveSize2D();
            size.setCx(cmToEMUs(width));
            size.setCy(cmToEMUs(height));
            inline.setExtent(size);
            inline.getGraphic().getGraphicData().getPic().getSpPr().getXfrm().setExt(size);
            // 保存运行数据
            Drawing drawing = FACTORY.createDrawing();
            drawing.getAnchorOrInline().add(inline);
            p.getContent().add(drawing);
        } catch (Exception e) {
            throw new IOException(e);
        }
    }

    /** {@inheritDoc} */
    @Override
    public void addSection(Section section) {
        if (!(section instanceof Docx4jSection))
            throw new IllegalArgumentException("section is not Docx4jSection");
        this.body.getContent().add(((Docx4jSection) section).getSectionOfDocx4j());
    }

    /** {@inheritDoc} */
    @Override
    public void updateTableOfContent() {
        try {
            TocGenerator generator = new TocGenerator(this.wml);
            generator.updateToc(true);
        } catch (TocException e) {
            e.printStackTrace();
        }
    }

    /** {@inheritDoc} */
    @Override
    public void mergeWith(Document document) {
        throw new UnsupportedOperationException("mergeWith is not yet supported");
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

    /** {@inheritDoc} */
    @Override
    public void saveAsHTML() throws IOException {
        throw new UnsupportedOperationException("save as HTMl is not yet supported");
    }

    private Docx4jDocument(WordprocessingMLPackage wml) {
        this.wml = wml;
        this.body = wml.getMainDocumentPart();
    }

    /**
     * 递归访问文档
     * @param content 访问内容
     * @param visitor 访问者
     * @return 访问结果
     */
    private DocumentVisitResult walk(Object content, DocumentVisitor visitor) {
        if (content instanceof ContentAccessor) {
            for (Object obj : ((ContentAccessor) content).getContent()) {
                DocumentVisitResult result = walk(obj, visitor);
                if (result == DocumentVisitResult.TERMINATE) return result;
            }
        } else if (content instanceof JAXBElement) {
            return walk(((JAXBElement<?>) content).getValue(), visitor);
        } else if (content instanceof Text && ((Text) content).getParent() instanceof R) {
            Text text = (Text) content;
            R r = (R) text.getParent();
            return visitor.visitText(new Docx4jText(r, text.getValue()));
        }
        return DocumentVisitResult.CONTINUE;
    }
}
