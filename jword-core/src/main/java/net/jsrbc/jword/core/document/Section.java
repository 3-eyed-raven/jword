package net.jsrbc.jword.core.document;

import net.jsrbc.jword.core.document.enums.HeaderFooterType;
import net.jsrbc.jword.core.document.enums.PageOrientation;
import net.jsrbc.jword.core.document.enums.PageSize;
import net.jsrbc.jword.core.document.enums.SectionType;

/**
 * 分节符
 * @author ZZZ on 2021/1/29 19:47
 * @version 1.0
 */
public interface Section {

    /**
     * 设置分节符类型
     * @param sectionType 分节符类型
     */
    void setType(SectionType sectionType);

    /**
     * 设置该节页面尺寸
     * @param pageSize 页面尺寸
     */
    void setPageSize(PageSize pageSize);

    /**
     * 设置页面方向
     * @param orientation 页面方向
     */
    void setPageOrientation(PageOrientation orientation);

    /**
     * 设置页边距，单位：cm
     * @param top 上边距，单位：cm
     * @param right 右边距，单位：cm
     * @param bottom 下边距，单位：cm
     * @param left 左边距，单位：cm
     */
    void setPageMargin(double top, double right, double bottom, double left);

    /**
     * 设置页眉距边界距离
     * @param headerMargin 页眉距边界距离，单位：cm
     */
    void setHeaderMargin(double headerMargin);

    /**
     * 设置页脚距边界距离
     * @param footerMargin 页脚距边界距离，单位：cm
     */
    void setFooterMargin(double footerMargin);

    /**
     * 添加页眉引用
     * @param id 页眉ID
     * @param headerFooterType 页眉类型
     */
    void addHeaderReference(String id, HeaderFooterType headerFooterType);

    /**
     * 添加页脚引用
     * @param id 页脚ID
     * @param headerFooterType 页脚类型
     */
    void addFooterReference(String id, HeaderFooterType headerFooterType);
}
