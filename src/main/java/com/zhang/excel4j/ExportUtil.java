package com.zhang.excel4j;

import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.handler.ColumnHandler;
import com.zhang.excel4j.handler.ExcelHandler;
import com.zhang.excel4j.handler.TemplateHandler;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

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
     * 基于注解，导出Excel数据到本地地址
     *
     * @param data     数据
     * @param clazz    处理对象
     * @param filePath 导出文件路径（包含后缀）
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
        ExcelHandler.exportWorkbook(data, clazz, null, null, workbookType).write(fos);
    }

    /**
     * 基于注解，导出Excel数据到输出流
     *
     * @param os           输出流
     * @param data         数据
     * @param clazz        处理对象
     * @param sheetName    sheet名称
     * @param groupName    分组名称
     * @param workbookType 工作簿类型
     * @throws Exception 异常
     */
    public void exportList2Excel(OutputStream os, List<?> data, Class<?> clazz, String sheetName, String groupName, WorkbookType workbookType) throws Exception {
        ExcelHandler.exportWorkbook(data, clazz, sheetName, groupName, workbookType).write(os);
    }

    /**
     * 基于注解，获取导出的Excel工作簿对象
     *
     * @param data         数据
     * @param clazz        处理对象
     * @param sheetName    sheet名称
     * @param groupName    分组名称
     * @param workbookType 工作簿类型
     * @return 工作簿对象
     * @throws Exception 异常
     */
    public Workbook getExportWorkbook(List<?> data, Class<?> clazz, String sheetName, String groupName, WorkbookType workbookType) throws Exception {
        return ExcelHandler.exportWorkbook(data, clazz, sheetName, groupName, workbookType);
    }

    /**
     * 无模板注解，导出Excel数据到本地地址
     *
     * @param data     数据
     * @param header   表头列表
     * @param filePath 导出文件路径（包含后缀）
     * @throws Exception 异常
     */
    public void exportList2Excel(List<?> data, List<String> header, String filePath) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        if (workbookType == null) {
            return;
        }
        ExcelHandler.exportWorkbook(data, header, null, workbookType).write(fos);
    }

    /**
     * 无模板注解，导出Excel数据到输出流
     *
     * @param os           输出流
     * @param data         数据
     * @param header       表头列表
     * @param sheetName    sheet名称
     * @param workbookType 工作簿类型
     * @throws Exception 异常
     */
    public void exportList2Excel(OutputStream os, List<?> data, List<String> header, String sheetName, WorkbookType workbookType) throws Exception {
        ExcelHandler.exportWorkbook(data, header, sheetName, workbookType).write(os);
    }

    /**
     * 无模板注解，获取导出的Excel工作簿对象
     *
     * @param data         数据
     * @param header       表头列表
     * @param sheetName    sheet名称
     * @param workbookType 工作簿类型
     * @return 工作簿对象
     */
    public Workbook getExportWorkbook(List<?> data, List<String> header, String sheetName, WorkbookType workbookType) {
        return ExcelHandler.exportWorkbook(data, header, sheetName, workbookType);
    }

    /**
     * 基于模板和注解，导出Excel数据到本地地址
     *
     * @param filePath      导出文件路径（包含后缀）
     * @param is            模板输入流
     * @param sheetIndex    模板sheet索引
     * @param data          表数据
     * @param extendData    额外数据
     * @param clazz         处理对象
     * @param groupName     分组名称
     * @param isWriteHeader 是否插入标题行
     * @param sheetName     sheet名称
     * @param isCopySheet   是否复制sheet
     * @throws Exception 异常
     */
    public void exportList2Excel(String filePath, InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                 String groupName, boolean isWriteHeader, String sheetName, boolean isCopySheet) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        if (workbookType == null) {
            return;
        }
        TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, clazz, groupName, isWriteHeader, sheetName, isCopySheet).getWorkbook().write(fos);
    }

    /**
     * 基于模板和注解，导出Excel数据到输出流
     *
     * @param os            输出流
     * @param is            模板输入流
     * @param sheetIndex    模板sheet索引
     * @param data          表数据
     * @param extendData    额外数据
     * @param clazz         处理对象
     * @param groupName     分组名称
     * @param isWriteHeader 是否插入标题行
     * @param sheetName     sheet名称
     * @param isCopySheet   是否复制sheet
     * @throws Exception 异常
     */
    public void exportList2Excel(OutputStream os, InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                 String groupName, boolean isWriteHeader, String sheetName, boolean isCopySheet) throws Exception {
        TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, clazz, groupName, isWriteHeader, sheetName, isCopySheet).getWorkbook().write(os);
    }

    /**
     * 基于模板和注解，获取导出的Excel工作簿对象
     *
     * @param is            模板输入流
     * @param sheetIndex    模板sheet索引
     * @param data          表数据
     * @param extendData    额外数据
     * @param clazz         处理对象
     * @param groupName     分组名称
     * @param isWriteHeader 是否插入标题行
     * @param sheetName     sheet名称
     * @param isCopySheet   是否复制sheet
     * @return 工作簿对象
     * @throws Exception 异常
     */
    public Workbook getExportWorkbook(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                      String groupName, boolean isWriteHeader, String sheetName, boolean isCopySheet) throws Exception {
        return TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, clazz, groupName, isWriteHeader, sheetName, isCopySheet).getWorkbook();
    }

    /**
     * 基于模板，导出Excel数据到本地地址
     *
     * @param filePath    导出文件路径（包含后缀）
     * @param is          模板输入流
     * @param sheetIndex  模板sheet索引
     * @param data        表数据
     * @param extendData  额外数据
     * @param sheetName   sheet名称
     * @param isCopySheet 是否复制sheet
     * @throws Exception 异常
     */
    public void exportList2Excel(String filePath, InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName, boolean isCopySheet) throws Exception {
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        // 工作簿类型
        WorkbookType workbookType = ColumnHandler.getWorkbookTypeByFilePath(filePath);
        if (workbookType == null) {
            return;
        }
        TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, sheetName, isCopySheet).getWorkbook().write(fos);
    }

    /**
     * 基于模板，导出Excel数据到输出流
     *
     * @param os          输出流
     * @param is          模板输入流
     * @param sheetIndex  模板sheet索引
     * @param data        表数据
     * @param extendData  额外数据
     * @param sheetName   sheet名称
     * @param isCopySheet 是否复制sheet
     * @throws Exception 异常
     */
    public void exportList2Excel(OutputStream os, InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName, boolean isCopySheet) throws Exception {
        TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, sheetName, isCopySheet).getWorkbook().write(os);
    }

    /**
     * 基于模板，获取导出的Excel工作簿对象
     *
     * @param is          模板输入流
     * @param sheetIndex  模板sheet索引
     * @param data        表数据
     * @param extendData  额外数据
     * @param sheetName   sheet名称
     * @param isCopySheet 是否复制sheet
     * @return 工作簿对象
     * @throws Exception 异常
     */
    public Workbook getExportWorkbook(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName, boolean isCopySheet) throws Exception {
        return TemplateHandler.exportExcelTemplate(is, sheetIndex, data, extendData, sheetName, isCopySheet).getWorkbook();
    }
}
