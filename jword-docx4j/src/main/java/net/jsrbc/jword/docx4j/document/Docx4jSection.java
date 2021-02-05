package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.Section;
import net.jsrbc.jword.core.document.enums.HeaderFooterType;
import net.jsrbc.jword.core.document.enums.PageOrientation;
import net.jsrbc.jword.core.document.enums.PageSize;
import net.jsrbc.jword.core.document.enums.SectionType;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.math.BigInteger;

import static net.jsrbc.jword.docx4j.util.UnitConverter.*;

/**
 * DOCX4J分节符
 * @author ZZZ on 2021/1/30 11:16
 * @version 1.0
 */
public class Docx4jSection implements Section {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    /** 分节符 */
    private final SectPr sectPr = FACTORY.createSectPr();

    /** {@inheritDoc} */
    @Override
    public void setType(SectionType sectionType) {
        SectPr.Type type = FACTORY.createSectPrType();
        type.setVal(sectionType.getValue());
        this.sectPr.setType(type);
    }

    /** {@inheritDoc} */
    @Override
    public void setPageSize(PageSize pageSize) {
        SectPr.PgSz pgSz = getPgSz();
        pgSz.setCode(BigInteger.valueOf(pageSize.getCode()));
        pgSz.setW(BigInteger.valueOf(cmToTwips(pageSize.getWidth())));
        pgSz.setH(BigInteger.valueOf(cmToTwips(pageSize.getHeight())));
    }

    /** {@inheritDoc} */
    @Override
    public void setPageOrientation(PageOrientation orientation) {
        SectPr.PgSz pgSz = getPgSz();
        pgSz.setOrient(STPageOrientation.valueOf(orientation.name()));
    }

    /** {@inheritDoc} */
    @Override
    public void setPageMargin(double top, double right, double bottom, double left) {
        SectPr.PgMar pgMar = getPgMar();
        pgMar.setTop(BigInteger.valueOf(cmToTwips(top)));
        pgMar.setRight(BigInteger.valueOf(cmToTwips(right)));
        pgMar.setBottom(BigInteger.valueOf(cmToTwips(bottom)));
        pgMar.setLeft(BigInteger.valueOf(cmToTwips(left)));
    }

    /** {@inheritDoc} */
    @Override
    public void setHeaderMargin(double headerMargin) {
        SectPr.PgMar pgMar = getPgMar();
        pgMar.setHeader(BigInteger.valueOf(cmToTwips(headerMargin)));
    }

    /** {@inheritDoc} */
    @Override
    public void setFooterMargin(double footerMargin) {
        SectPr.PgMar pgMar = getPgMar();
        pgMar.setFooter(BigInteger.valueOf(cmToTwips(footerMargin)));
    }

    /** {@inheritDoc} */
    @Override
    public void addHeaderReference(String id, HeaderFooterType headerFooterType) {
        HeaderReference hr = FACTORY.createHeaderReference();
        hr.setId(id);
        hr.setType(HdrFtrRef.valueOf(headerFooterType.name()));
        this.sectPr.getEGHdrFtrReferences().add(hr);
    }

    /** {@inheritDoc} */
    @Override
    public void addFooterReference(String id, HeaderFooterType headerFooterType) {
        FooterReference fr = FACTORY.createFooterReference();
        fr.setId(id);
        fr.setType(HdrFtrRef.valueOf(headerFooterType.name()));
        this.sectPr.getEGHdrFtrReferences().add(fr);
    }

    /**
     * 返回Docx4j分节符
     * @return 分节符
     */
    public SectPr getSectionOfDocx4j() {
        return this.sectPr;
    }

    /**
     * 获取页面尺寸对象
     * @return 页面尺寸对象
     */
    private SectPr.PgSz getPgSz() {
        if (sectPr.getPgSz() == null) {
            SectPr.PgSz pgSz = FACTORY.createSectPrPgSz();
            this.sectPr.setPgSz(pgSz);
            return pgSz;
        }
        return this.sectPr.getPgSz();
    }

    /**
     * 获取页面尺寸对象
     * @return 页面尺寸对象
     */
    private SectPr.PgMar getPgMar() {
        if (sectPr.getPgMar() == null) {
            SectPr.PgMar pgMar = FACTORY.createSectPrPgMar();
            this.sectPr.setPgMar(pgMar);
            return pgMar;
        }
        return this.sectPr.getPgMar();
    }
}
