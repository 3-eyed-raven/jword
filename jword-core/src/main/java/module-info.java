module net.jsrbc.jword.core {
    requires org.reactivestreams;
    requires reactor.core;

    exports net.jsrbc.jword.core;
    exports net.jsrbc.jword.core.api;
    exports net.jsrbc.jword.core.document;
    exports net.jsrbc.jword.core.document.enums;
    exports net.jsrbc.jword.core.factory;
}