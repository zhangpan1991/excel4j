package com.zhang.excel4j.model;

import java.util.List;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/4/4 16:19
 */
public class ExportData {

    /**
     * 模板sheet索引
     */
    private int sheetIndex;

    /**
     * 表数据
     */
    private List<?> data;

    /**
     * 额外数据
     */
    private Map<String, Object> extendData;

    /**
     * 处理对象
     */
    private Class<?> clazz;

    /**
     * 分组名称
     */
    private String groupName;

    /**
     * 是否插入标题行
     */
    private boolean writeHeader;

    /**
     * sheet名称
     */
    private String sheetName;

    public ExportData() {
    }

    public ExportData(int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName) {
        this.sheetIndex = sheetIndex;
        this.data = data;
        this.extendData = extendData;
        this.sheetName = sheetName;
    }

    public ExportData(int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz, String groupName, boolean writeHeader, String sheetName) {
        this.sheetIndex = sheetIndex;
        this.data = data;
        this.extendData = extendData;
        this.clazz = clazz;
        this.groupName = groupName;
        this.writeHeader = writeHeader;
        this.sheetName = sheetName;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }

    public Map<String, Object> getExtendData() {
        return extendData;
    }

    public void setExtendData(Map<String, Object> extendData) {
        this.extendData = extendData;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public String getGroupName() {
        return groupName;
    }

    public void setGroupName(String groupName) {
        this.groupName = groupName;
    }

    public boolean isWriteHeader() {
        return writeHeader;
    }

    public void setWriteHeader(boolean writeHeader) {
        this.writeHeader = writeHeader;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }
}
