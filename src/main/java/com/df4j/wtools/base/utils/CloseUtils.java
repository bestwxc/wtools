package com.df4j.wtools.base.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;

public class CloseUtils {

    private static Logger logger = LoggerFactory.getLogger(CloseUtils.class);

    public static void close(Closeable closeable) {
        if (closeable == null) {
            return;
        }
        try {
            closeable.close();
        } catch (Exception e) {
            logger.warn("关闭资源错误,type: " + closeable.getClass().getName(), e);
        }
    }
}
