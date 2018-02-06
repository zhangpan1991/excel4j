package com.zhang.excel4j;

import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.handler.ColumnHandler;
import com.zhang.excel4j.handler.ExcelHandler;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * author : zhangpan
 * date : 2018/1/25 19:21
 */
public class ExportUtil {

    /**
     * 单例模式
     */
    private volatile static ExportUtil exportUtil;

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
    public void exportList2Excel(List<?> data, Class<?> clazz, String filePath) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        if (workbookType == null) {
            return;
        }
        exportList2Excel(fos, data, clazz, null, null, workbookType);
    }

    public void exportList2Excel(OutputStream os, List<?> data, Class<?> clazz, String sheetName, String groupName, WorkbookType workbookType) throws Exception {
        ExcelHandler.exportWorkbook(data, clazz, sheetName, groupName, workbookType).write(os);
    }

    public void exportList2Excel(List<?> data, List<String> header, String filePath) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        if (workbookType == null) {
            return;
        }
        exportList2Excel(fos, data, header, null, workbookType);
    }

    public void exportList2Excel(OutputStream os, List<?> data, List<String> header, String sheetName, WorkbookType workbookType) throws Exception {
        ExcelHandler.exportWorkbook(data, header, sheetName, workbookType).write(os);
    }
}
