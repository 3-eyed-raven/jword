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

    /** 单元格对象 */
    private final Tc tc = FACTORY.createTc();

    /** {@inheritDoc} */
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

    /** {@inheritDoc} */
    @Override
    public void setGridSpan(int span) {
        TcPrInner.GridSpan gridSpan = FACTORY.createTcPrInnerGridSpan();
        gridSpan.setVal(BigInteger.valueOf(span));
        getTcPr().setGridSpan(gridSpan);
    }

    /** {@inheritDoc} */
    @Override
    public void setVerticalMergeType(VerticalMergeType verticalMergeType) {
        TcPrInner.VMerge vMerge = FACTORY.createTcPrInnerVMerge();
        vMerge.setVal(verticalMergeType.getValue());
        getTcPr().setVMerge(vMerge);
    }

    /** {@inheritDoc} */
    @Override
    public void setVerticalAlignType(VerticalAlignType verticalAlignType) {
        CTVerticalJc vAlign = FACTORY.createCTVerticalJc();
        vAlign.setVal(STVerticalJc.valueOf(verticalAlignType.name()));
        getTcPr().setVAlign(vAlign);
    }

    /** {@inheritDoc} */
    @Override
    public void addParagraph(Paragraph paragraph) {
        if (!(paragraph instanceof Docx4jParagraph))
            throw new IllegalArgumentException("paragraph is not Docx4jParagraph");
        this.tc.getContent().add(((Docx4jParagraph) paragraph).getParagraphOfDocx4j());
    }

    /**
     * 获取Docx4j对应的表格单元格对象
     */
    public Tc getTableCellOfDocx4j() {
        return this.tc;
    }

    /**
     * 获取表格单元格，不存在则创建
     */
    private TcPr getTcPr() {
        if (this.tc.getTcPr() == null) {
            TcPr tcPr = FACTORY.createTcPr();
            this.tc.setTcPr(tcPr);
            return tcPr;
        }
        return this.tc.getTcPr();
    }
}
