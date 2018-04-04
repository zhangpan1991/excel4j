package com.zhang.excel4j.handler;

import com.zhang.excel4j.common.TemplateConstant;
import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.model.ExcelHeader;
import com.zhang.excel4j.model.ExportData;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/2/2 18:39
 */
public class ExcelTemplate {

    /**
     * 模板工作簿
     */
    private Workbook template;

    /**
     * 当前工作簿
     */
    private Workbook workbook;

    /**
     * 模板sheet表
     */
    private Sheet tempSheet;

    /**
     * 当前sheet
     */
    private Sheet sheet;

    /**
     * 当前行
     */
    private Row row;

    /**
     * 当前列下标
     */
    private int currentCol;

    /**
     * 当前行下标
     */
    private int currentRow;

    /**
     * 序号
     */
    private int serialNumber;

    /**
     * 序号列坐标
     */
    private int serialNumberCol = -1;

    /**
     * 自定义数据映射关系
     */
    private Map<Integer, String> headerMap = new HashMap<>();

    /**
     * 默认样式
     */
    private CellStyle defaultStyle;

    /**
     * 指定行样式
     */
    private Map<Integer, CellStyle> appointLineStyle = new HashMap<>();

    /**
     * 单数行样式
     */
    private CellStyle singleLineStyle;

    /**
     * 双数行样式
     */
    private CellStyle doubleLineStyle;

    /**
     * 数据的初始化列数
     */
    private int initCol;

    /**
     * 数据的初始化行数
     */
    private int initRow;

    /**
     * 最后一行的数据
     */
    private int lastRow;

    /**
     * 默认行高
     */
    private float rowHeight;

    /**
     * 加载Excel模板
     *
     * @param is           模板输入流
     * @param workbookType 生成工作簿类型
     * @return Excel模板对象
     * @throws IOException            异常
     * @throws InvalidFormatException 异常
     */
    public static ExcelTemplate loadTemplate(InputStream is, WorkbookType workbookType) throws IOException, InvalidFormatException {
        Workbook template = WorkbookFactory.create(is);
        Workbook workbook = ExcelHandler.createWorkbook(workbookType);
        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setTemplate(template);
        excelTemplate.setWorkbook(workbook);
        return excelTemplate;
    }

    /**
     * 加载模板Sheet
     *
     * @param sheetIndex Sheet索引
     * @param extendData 额外数据
     */
    private void loadTempSheet(int sheetIndex, Map<String, Object> extendData) {
        // sheet模板
        Sheet tempSheet = this.template.getSheetAt(sheetIndex);
        this.tempSheet = tempSheet;
        this.lastRow = tempSheet.getLastRowNum();

        for (Row row : tempSheet) {
            for (Cell cell : row) {
                if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                    String value = cell.getStringCellValue().trim();
                    // 替换额外数据
                    if (value.startsWith(TemplateConstant.EXTEND_DATA_SIGN)) {
                        if (extendData != null && extendData.containsKey(value.substring(1))) {
                            cell.setCellValue(extendData.get(value.substring(1)).toString());
                        }
                        continue;
                    }
                    // 自定义数据映射表头
                    if (value.startsWith(TemplateConstant.DATA_HEADER_SIGN)) {
                        this.headerMap.put(cell.getColumnIndex(), value.substring(1));
                        continue;
                    }
                    // 获取模板数据和样式
                    switch (value) {
                        // 序号列
                        case TemplateConstant.SERIAL_NUMBER:
                            this.serialNumberCol = cell.getColumnIndex();
                            SheetHandler.clearCell(cell);
                            break;
                        // 数据列
                        case TemplateConstant.DATA_INDEX:
                            this.initCol = cell.getColumnIndex();
                            this.initRow = cell.getRowIndex();
                            this.currentCol = cell.getColumnIndex();
                            this.currentRow = cell.getRowIndex();
                            this.rowHeight = row.getHeightInPoints();
                            SheetHandler.clearCell(cell);
                            break;
                        // 默认行样式
                        case TemplateConstant.DEFAULT_STYLE:
                            this.defaultStyle = cell.getCellStyle();
                            SheetHandler.clearCell(cell);
                            break;
                        // 单行样式
                        case TemplateConstant.APPOINT_LINE_STYLE:
                            this.appointLineStyle.put(cell.getRowIndex(), cell.getCellStyle());
                            SheetHandler.clearCell(cell);
                            break;
                        // 单数行样式
                        case TemplateConstant.SINGLE_LINE_STYLE:
                            this.singleLineStyle = cell.getCellStyle();
                            SheetHandler.clearCell(cell);
                            break;
                        // 双数行样式
                        case TemplateConstant.DOUBLE_LINE_STYLE:
                            this.doubleLineStyle = cell.getCellStyle();
                            SheetHandler.clearCell(cell);
                            break;
                    }
                }
            }
        }
    }

    /**
     * 装载数据
     *
     * @param sheetIndex 模板sheet索引
     * @param data       表数据
     * @param extendData 额外数据
     * @param sheetName  sheet名称
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate loadData(int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName) throws Exception {
        return loadData(new ExportData(sheetIndex, data, extendData, sheetName));
    }

    /**
     * 装载数据
     *
     * @param sheetIndex  模板sheet索引
     * @param data        表数据
     * @param extendData  额外数据
     * @param clazz       处理对象
     * @param groupName   分组名称
     * @param writeHeader 是否插入标题行
     * @param sheetName   sheet名称
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate loadData(int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz, String groupName, boolean writeHeader, String sheetName) throws Exception {
        return loadData(new ExportData(sheetIndex, data, extendData, clazz, groupName, writeHeader, sheetName));
    }

    /**
     * 装载数据
     *
     * @param exportDataList 导出数据集合
     * @return Excel模板
     * @throws Exception 异常
     */
    public ExcelTemplate loadData(List<ExportData> exportDataList) throws Exception {
        for (ExportData exportData : exportDataList) {
            loadData(exportData);
        }
        return this;
    }

    /**
     * 装载数据
     *
     * @param exportData 导出数据对象
     * @return Excel模板
     * @throws Exception 异常
     */
    private ExcelTemplate loadData(ExportData exportData) throws Exception {
        // 加载Sheet模板
        this.loadTempSheet(exportData.getSheetIndex(), exportData.getExtendData());
        // 创建Sheet
        this.createSheet(exportData.getSheetName());
        Class clazz = exportData.getClazz();
        if (clazz != null) {
            // 获取表头列表
            String groupName = exportData.getGroupName();
            List<ExcelHeader> headers;
            if (groupName != null && !"".equals(groupName)) {
                headers = ColumnHandler.getExcelHeaderList(groupName, clazz);
            } else {
                headers = ColumnHandler.getExcelHeaderList(clazz);
            }
            if (exportData.isWriteHeader()) {
                this.createRow();
                this.createSerialNumber(true);
                for (ExcelHeader header : headers) {
                    this.createCell(header.getTitle());
                }
            }
            for (Object object : exportData.getData()) {
                this.createRow();
                this.createSerialNumber(false);
                for (ExcelHeader header : headers) {
                    if (object instanceof Map) {
                        // Map数据
                        this.createCell(ColumnHandler.getValueByMap((Map) object, header.getFiled(), header.getConverter()));
                    } else {
                        // 处理对象数据
                        this.createCell(ColumnHandler.getValueByAttribute(object, header.getFiled(), header.getConverter()));
                    }
                }
            }
        } else {
            for (Object object : exportData.getData()) {
                this.createRow();
                this.createSerialNumber(false);
                for (Map.Entry<Integer, String> entry : this.getHeaderMap().entrySet()) {
                    if (object instanceof Map) {
                        // Map数据
                        this.createCell(ColumnHandler.getValueByMap((Map) object, entry.getValue(), null), entry.getKey());
                    } else {
                        // 处理对象数据
                        this.createCell(ColumnHandler.getValueByAttribute(object, entry.getValue(), null), entry.getKey());
                    }
                }
            }
        }
        return this;
    }

    /**
     * 新增Sheet
     * @param sheetName sheet名称
     */
    public void createSheet(String sheetName) {
        // 复制模板
        Sheet sheet = this.workbook.createSheet();
        SheetHandler.copySheet(tempSheet, sheet, 0, tempSheet.getLastRowNum());
        // 重置序列号
        this.serialNumber = 1;
        // 修改sheet名称
        if (sheetName != null && !"".equals(sheetName)) {
            int index = this.workbook.getSheetIndex(sheet);
            this.workbook.setSheetName(index, sheetName);
        }
        this.sheet = sheet;
    }

    /**
     * 新增行
     */
    private void createRow() {
        if (this.lastRow > this.currentRow && this.currentRow != this.initRow) {
            this.sheet.shiftRows(this.currentRow, this.lastRow, 1, true, true);
            this.lastRow++;
        }
        this.row = this.sheet.createRow(this.currentRow);
        this.row.setHeightInPoints(this.rowHeight);
        this.currentRow++;
        this.currentCol = this.initCol;
    }

    /**
     * 新增单元格
     *
     * @param value 数据
     */
    private void createCell(Object value) {
        Cell cell = this.row.createCell(currentCol);
        // 设置单元格样式
        setCellStyle(cell);
        cell.setCellValue(value.toString());
        this.currentCol++;
    }

    /**
     * 新增单元格
     *
     * @param value 数据
     * @param col   列
     */
    private void createCell(Object value, int col) {
        Cell cell = this.row.createCell(col);
        // 设置单元格样式
        setCellStyle(cell);
        cell.setCellValue(value.toString());
    }

    /**
     * 创建序号列
     *
     * @param isHeader 是否表头
     */
    private void createSerialNumber(boolean isHeader) {
        if (this.serialNumberCol < 0) {
            return;
        }
        Cell cell = this.row.createCell(this.serialNumberCol);
        setCellStyle(cell);
        if (isHeader) {
            cell.setCellValue("序号");
        } else {
            cell.setCellValue(this.serialNumber++);
        }
    }

    /**
     * 设置单元格样式
     * 优先级：单行样式 > 单数行样式 = 双数行样式 > 默认样式
     *
     * @param cell 单元格
     */
    private void setCellStyle(Cell cell) {
        if (this.appointLineStyle.containsKey(cell.getRowIndex())) {
            cell.getCellStyle().cloneStyleFrom(this.appointLineStyle.get(cell.getRowIndex()));
            return;
        }
        if (null != this.singleLineStyle && (cell.getRowIndex() % 2 != 0)) {
            cell.getCellStyle().cloneStyleFrom(this.singleLineStyle);
            return;
        }
        if (null != this.doubleLineStyle && (cell.getRowIndex() % 2 == 0)) {
            cell.getCellStyle().cloneStyleFrom(this.doubleLineStyle);
            return;
        }
        if (null != this.defaultStyle)
            cell.getCellStyle().cloneStyleFrom(this.defaultStyle);
    }

    /**
     * 导出Excel数据到本地地址
     *
     * @param filePath 导出文件路径（包含后缀）
     * @throws IOException IO异常
     */
    public void export(String filePath) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            this.workbook.write(fos);
        }
    }

    /**
     * 导出Excel数据到输出流
     *
     * @param os 输出流
     * @throws IOException IO异常
     */
    public void export(OutputStream os) throws IOException {
        this.workbook.write(os);
    }

    public Workbook getTemplate() {
        return template;
    }

    public void setTemplate(Workbook template) {
        this.template = template;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getTempSheet() {
        return tempSheet;
    }

    public void setTempSheet(Sheet tempSheet) {
        this.tempSheet = tempSheet;
    }

    public int getSerialNumber() {
        return serialNumber;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public int getCurrentCol() {
        return currentCol;
    }

    public void setCurrentCol(int currentCol) {
        this.currentCol = currentCol;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public void setCurrentRow(int currentRow) {
        this.currentRow = currentRow;
    }

    public void setSerialNumber(int serialNumber) {
        this.serialNumber = serialNumber;
    }

    public int getSerialNumberCol() {
        return serialNumberCol;
    }

    public void setSerialNumberCol(int serialNumberCol) {
        this.serialNumberCol = serialNumberCol;
    }

    public Map<Integer, String> getHeaderMap() {
        return headerMap;
    }

    public void setHeaderMap(Map<Integer, String> headerMap) {
        this.headerMap = headerMap;
    }

    public CellStyle getDefaultStyle() {
        return defaultStyle;
    }

    public void setDefaultStyle(CellStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public Map<Integer, CellStyle> getAppointLineStyle() {
        return appointLineStyle;
    }

    public void setAppointLineStyle(Map<Integer, CellStyle> appointLineStyle) {
        this.appointLineStyle = appointLineStyle;
    }

    public CellStyle getSingleLineStyle() {
        return singleLineStyle;
    }

    public void setSingleLineStyle(CellStyle singleLineStyle) {
        this.singleLineStyle = singleLineStyle;
    }

    public CellStyle getDoubleLineStyle() {
        return doubleLineStyle;
    }

    public void setDoubleLineStyle(CellStyle doubleLineStyle) {
        this.doubleLineStyle = doubleLineStyle;
    }

    public int getInitCol() {
        return initCol;
    }

    public void setInitCol(int initCol) {
        this.initCol = initCol;
    }

    public int getInitRow() {
        return initRow;
    }

    public void setInitRow(int initRow) {
        this.initRow = initRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public float getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(float rowHeight) {
        this.rowHeight = rowHeight;
    }
}
