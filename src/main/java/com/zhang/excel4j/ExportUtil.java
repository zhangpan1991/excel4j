package com.zhang.excel4j;

import com.zhang.excel4j.common.WorkbookType;
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

    /**
     * 导出数据到本地地址
     *
     * @param data     数据
     * @param clazz    处理对象
     * @param filePath 文件路径（包含后缀）
     * @throws Exception 异常
     */
    public void exportObjects2Excel(List<?> data, Class clazz, String filePath) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        ExcelHandler.exportWorkbookWithAnnotation(data, clazz, null, null, true, workbookType).write(fos);
    }
}
