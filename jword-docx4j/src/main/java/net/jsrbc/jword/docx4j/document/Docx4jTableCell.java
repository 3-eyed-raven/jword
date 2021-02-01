package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.TableCell;
import net.jsrbc.jword.core.document.enums.TableWidthType;
import net.jsrbc.jword.core.document.enums.VerticalAlignType;
import net.jsrbc.jword.core.document.enums.VerticalMergeType;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.Tc;

/**
 * 表格单元格
 * @author ZZZ on 2021/2/1 22:01
 * @version 1.0
 */
public class Docx4jTableCell implements TableCell {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    private final Tc tc = FACTORY.createTc();

    @Override
    public void setCellWidth(int width, TableWidthType widthType) {

    }

    @Override
    public void setGridSpan(int span) {

    }

    @Override
    public void setVerticalMergeType(VerticalMergeType verticalMergeType) {

    }

    @Override
    public void setVerticalAlignType(VerticalAlignType verticalAlignType) {

    }

    @Override
    public void addParagraph(Paragraph paragraph) {

    }

    public Tc getTableCellOfDocx4j() {
        return this.tc;
    }
}
