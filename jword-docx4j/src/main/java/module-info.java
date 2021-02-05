module net.jsrbc.jword.docx4j {
    requires org.docx4j.core;
    requires org.docx4j.JAXB_ReferenceImpl;
    requires net.jsrbc.jword.core;

    exports net.jsrbc.jword.docx4j.document;
    exports net.jsrbc.jword.docx4j.factory;
}