package net.jsrbc.jword.core.api;

import java.io.InputStream;
import java.nio.file.Path;

/**
 * Jword加载器
 * @author ZZZ on 2021/2/4 10:24
 * @version 1.0
 */
public interface JwordLoader {
    /**
     * 从指定路径加载Jword对象
     * @param templatePath 模板路径
     * @return Jword操作器
     */
    JwordOperator load(Path templatePath);

    /**
     * 从输入流中加载模板文件
     * @param in 输入流
     * @return Jword操作器
     */
    JwordOperator load(InputStream in);
}
