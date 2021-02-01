package net.jsrbc.jword.docx4j.document;

import net.jsrbc.jword.core.document.TableCell;
import net.jsrbc.jword.core.document.TableRow;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.Tr;

/**
 * 表格行
 * @author ZZZ on 2021/2/1 22:00
 * @version 1.0
 */
public class Docx4jTableRow implements TableRow {

    private final static ObjectFactory FACTORY = Context.getWmlObjectFactory();

    private final Tr tr = FACTORY.createTr();

    @Override
    public void addCell(TableCell cell) {
        if (!(cell instanceof Docx4jTableCell))
            throw new IllegalArgumentException("cell is not Docx4jTableCell");
        this.tr.getContent().add(((Docx4jTableCell) cell).getTableCellOfDocx4j());
    }

    public Tr getTableRowOfDocx4j() {
        return this.tr;
    }
}
