package net.jsrbc.jword.core.document.enums;

/**
 * 页眉页脚类型
 * 在文档的每节中，最多可以存在三种不同类型的页眉：
 * 1、首页页眉（FIRST）
 * 2、奇数页页眉(DEFAULT)
 * 3、偶数页页眉(EVEN)
 * @author ZZZ on 2021/1/30 10:25
 * @version 1.0
 */
public enum HeaderFooterType {
    /**
     * 首页页眉页脚
     * 如果没有为第一页指定 headerReference 但指定了 titlePg 元素，则从上一节继承首页页眉
     * 如果这是文档的第一节，则创建新的空白页眉
     * 如果未指定 titlePg 元素，则不显示首页页眉，并在其位置使用奇数页页眉
     */
    FIRST,

    /**
     * 偶数页页眉页脚
     * 如果没有为偶数页页眉指定 headerReference 但指定了 evenAndOddHeaders 元素，则从上一节继承偶数页页眉
     * 如果这是文档的第一节，则创建新的空白页眉
     * 如果未指定 evenAndOddHeaders 元素，则不显示偶数页页眉，并在其位置使用奇数页页眉
     */
    EVEN,

    /**
     * 默认
     *如果没有为奇数页页眉指定 headerReference ，则从上一节继承偶数页页眉；如果这是文档的第一节，则创建新的空白页眉
     */
    DEFAULT
}
