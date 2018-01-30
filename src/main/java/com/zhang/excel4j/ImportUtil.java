package com.zhang.excel4j;

/**
 * author : zhangpan
 * date : 2018/1/25 19:21
 */
public class ImportUtil {

    /**
     * 单例模式
     */
    private static ImportUtil importUtil;

    private ImportUtil() {
    }

    public static ImportUtil getInstance() {
        if (importUtil == null) {
            synchronized (ImportUtil.class) {
                if (importUtil == null) {
                    importUtil = new ImportUtil();
                }
            }
        }
        return importUtil;
    }
}
