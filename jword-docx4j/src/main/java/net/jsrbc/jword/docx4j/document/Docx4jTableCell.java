package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Paragraph;
import net.jsrbc.jword.core.document.TableCell;
import net.jsrbc.jword.core.document.enums.TableWidthType;
import net.jsrbc.jword.core.document.enums.VerticalAlignType;
import net.jsrbc.jword.core.document.enums.VerticalMergeType;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.math.BigInteger;

import static net.jsrbc.jword.docx4j.util.UnitConverter.cmToTwips;
import static net.jsrbc.jword.docx4j.util.UnitConverter.percentToFifthPct;

/**
 * 表格单元格
 * @author ZZZ on 2021/2/1 22:01
 * @version 1.0
 */
public class Docx4jTableCell implements TableCell {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    private final Tc tc = FACTORY.createTc();

    @Override
    public void setCellWidth(double width, TableWidthType type) {
        TblWidth tcW = FACTORY.createTblWidth();
        switch (type) {
            case DXA -> tcW.setW(BigInteger.valueOf(cmToTwips(width)));
            case PCT -> tcW.setW(BigInteger.valueOf(percentToFifthPct(width)));
            default -> tcW.setW(BigInteger.ZERO);
        }
        tcW.setType(type.getValue());
        getTcPr().setTcW(tcW);
    }

    @Override
    public void setGridSpan(int span) {
        TcPrInner.GridSpan gridSpan = FACTORY.createTcPrInnerGridSpan();
        gridSpan.setVal(BigInteger.valueOf(span));
        getTcPr().setGridSpan(gridSpan);
    }

    @Override
    public void setVerticalMergeType(VerticalMergeType verticalMergeType) {
        TcPrInner.VMerge vMerge = FACTORY.createTcPrInnerVMerge();
        vMerge.setVal(verticalMergeType.getValue());
        getTcPr().setVMerge(vMerge);
    }

    @Override
    public void setVerticalAlignType(VerticalAlignType verticalAlignType) {
        CTVerticalJc vAlign = FACTORY.createCTVerticalJc();
        vAlign.setVal(STVerticalJc.valueOf(verticalAlignType.name()));
        getTcPr().setVAlign(vAlign);
    }

    @Override
    public void addParagraph(Paragraph paragraph) {
        if (!(paragraph instanceof Docx4jParagraph))
            throw new IllegalArgumentException("paragraph is not Docx4jParagraph");
        this.tc.getContent().add(((Docx4jParagraph) paragraph).getParagraphOfDocx4j());
    }

    public Tc getTableCellOfDocx4j() {
        return this.tc;
    }

    private TcPr getTcPr() {
        if (this.tc.getTcPr() == null) {
            TcPr tcPr = FACTORY.createTcPr();
            this.tc.setTcPr(tcPr);
            return tcPr;
        }
        return this.tc.getTcPr();
    }
}
