package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Table;
import net.jsrbc.jword.core.document.TableRow;
import net.jsrbc.jword.core.document.enums.TableJustification;
import net.jsrbc.jword.core.document.enums.TableWidthType;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.math.BigInteger;

import static net.jsrbc.jword.docx4j.util.UnitConverter.*;

/**
 * DOCX4J表格
 * @author ZZZ on 2021/2/1 21:53
 * @version 1.0
 */
public class Docx4jTable implements Table {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    private final Tbl tbl = FACTORY.createTbl();

    /** {@inheritDoc} */
    @Override
    public void setStyle(String styleId) {
        CTTblPrBase.TblStyle style = FACTORY.createCTTblPrBaseTblStyle();
        style.setVal(styleId);
        getTblPr().setTblStyle(style);
    }

    /** {@inheritDoc} */
    @Override
    public void setWidth(double width, TableWidthType type) {
        TblWidth tblWidth = FACTORY.createTblWidth();
        switch (type) {
            case DXA:
                tblWidth.setW(BigInteger.valueOf(cmToTwips(width)));
                break;
            case PCT:
                tblWidth.setW(BigInteger.valueOf(percentToFifthPct(width)));
                break;
            default:
                tblWidth.setW(BigInteger.ZERO);
                break;
        }
        tblWidth.setType(type.getValue());
        getTblPr().setTblW(tblWidth);
    }

    /** {@inheritDoc} */
    @Override
    public void setJustification(TableJustification tableJustification) {
        Jc jc = FACTORY.createJc();
        jc.setVal(JcEnumeration.valueOf(tableJustification.name()));
        getTblPr().setJc(jc);
    }

    /** {@inheritDoc} */
    @Override
    public void addRow(TableRow row) {
        if (!(row instanceof Docx4jTableRow))
            throw new IllegalArgumentException("row is not Docx4jTableRow");
        this.tbl.getContent().add(((Docx4jTableRow) row).getTableRowOfDocx4j());
    }

    /**
     * 获取Docx4j对应的表格对象
     */
    public Tbl getTableOfDocx4j() {
        return this.tbl;
    }

    /**
     * 获取表格对象，不存在则创建
     */
    private TblPr getTblPr() {
        if (this.tbl.getTblPr() == null) {
            TblPr tblPr = FACTORY.createTblPr();
            tbl.setTblPr(tblPr);
            return tblPr;
        }
        return this.tbl.getTblPr();
    }
}
