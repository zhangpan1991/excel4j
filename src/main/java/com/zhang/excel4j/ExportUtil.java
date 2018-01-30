package com.zhang.excel4j;

import com.zhang.excel4j.handler.ColumnHandler;
import com.zhang.excel4j.handler.ExcelHandler;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/**
 * author : zhangpan
 * date : 2018/1/25 19:21
 */
public class ExportUtil {

    /**
     * 单例模式
     */
    private static ExportUtil exportUtil;

    private ExportUtil() {
    }

    public static ExportUtil getInstance() {
        if (exportUtil == null) {
            synchronized (ExportUtil.class) {
                if (exportUtil == null) {
                    exportUtil = new ExportUtil();
                }
            }
        }
        return exportUtil;
    }

    public void exportObjects2Excel(List<?> data, Class clazz, String filePath) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        ExcelHandler.exportWorkbookWithAnnotation(data, clazz, null, null, true, ColumnHandler.getWorkbookTypeByFilePath(filePath)).write(fos);
    }
}
