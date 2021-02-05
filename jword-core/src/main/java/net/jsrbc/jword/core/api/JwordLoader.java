package net.jsrbc.jword.core.api;

import java.nio.file.Path;

/**
 * Jword加载器
 * @author ZZZ on 2021/2/4 10:24
 * @version 1.0
 */
public interface JwordLoader {
    /**
     * 从模板加载Jword对象
     * @param templatePath 模板路径
     * @return Jword操作器
     */
    JwordOperator load(Path templatePath);
}
