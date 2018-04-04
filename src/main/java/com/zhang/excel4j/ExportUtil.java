package com.zhang.excel4j;

import com.zhang.excel4j.common.ExportType;
import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.handler.ExcelHandler;
import com.zhang.excel4j.handler.ExcelTemplate;
import com.zhang.excel4j.model.ExportData;
import com.zhang.excel4j.model.ExportModel;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.InputStream;
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
     * 基于注解，获取导出的Excel导出模型
     *
     * @param data         数据
     * @param clazz        处理对象
     * @param sheetName    sheet名称
     * @param groupName    分组名称
     * @param workbookType 工作簿类型
     * @return 导出模型
     * @throws Exception 异常
     */
    public ExportModel getExportModel(List<?> data, Class<?> clazz, String sheetName, String groupName, WorkbookType workbookType) throws Exception {
        Workbook workbook = ExcelHandler.exportWorkbook(data, clazz, sheetName, groupName, workbookType);
        return new ExportModel(ExportType.WORKBOOK, workbook);
    }

    /**
     * 无模板注解，获取导出的Excel导出模型
     *
     * @param data         数据
     * @param header       表头列表
     * @param sheetName    sheet名称
     * @param workbookType 工作簿类型
     * @return 导出模型
     */
    public ExportModel getExportModel(List<?> data, List<String> header, String sheetName, WorkbookType workbookType) {
        Workbook workbook = ExcelHandler.exportWorkbook(data, header, sheetName, workbookType);
        return new ExportModel(ExportType.WORKBOOK, workbook);
    }

    /**
     * 基于模板和注解，获取导出的Excel模板，装载数据
     *
     * @param is            模板输入流
     * @param sheetIndex    模板sheet索引
     * @param data          表数据
     * @param extendData    额外数据
     * @param clazz         处理对象
     * @param groupName     分组名称
     * @param isWriteHeader 是否插入标题行
     * @param sheetName     sheet名称
     * @param workbookType  工作簿类型
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate getExcelTemplate(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                        String groupName, boolean isWriteHeader, String sheetName, WorkbookType workbookType) throws Exception {
        sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
        return ExcelTemplate.loadTemplate(is, workbookType).loadData(sheetIndex, data, extendData, clazz, groupName, isWriteHeader, sheetName);
    }

    /**
     * 基于模板，获取导出的Excel模板，装载数据
     *
     * @param is           模板输入流
     * @param sheetIndex   模板sheet索引
     * @param data         表数据
     * @param extendData   额外数据
     * @param sheetName    sheet名称
     * @param workbookType 工作簿类型
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate getExcelTemplate(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName, WorkbookType workbookType) throws Exception {
        sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
        return ExcelTemplate.loadTemplate(is, workbookType).loadData(sheetIndex, data, extendData, sheetName);
    }

    /**
     * 基于模板，获取导出的Excel模板
     *
     * @param is           模板输入流
     * @param workbookType 工作簿类型
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate getExcelTemplate(InputStream is, WorkbookType workbookType) throws Exception {
        return ExcelTemplate.loadTemplate(is, workbookType);
    }

    /**
     * 基于模板，获取导出的Excel模板，装载数据
     *
     * @param is             模板输入流
     * @param workbookType   工作簿类型
     * @param exportDataList 导出数据集合
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate getExcelTemplate(InputStream is, WorkbookType workbookType, List<ExportData> exportDataList) throws Exception {
        return ExcelTemplate.loadTemplate(is, workbookType).loadData(exportDataList);
    }
}
