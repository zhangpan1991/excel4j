package com.zhang.excel4j.model;

import com.zhang.excel4j.common.ExportType;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;

/**
 * author : zhangpan
 * date : 2018/4/4 14:46
 */
public class ExportModel {

    /**
     * 导出类型
     */
    private ExportType exportType;

    /**
     * 导出的工作簿
     */
    private Workbook workbook;

    public ExportModel(ExportType exportType, Workbook workbook) {
        this.exportType = exportType;
        this.workbook = workbook;
    }

    /**
     * 导出Excel数据到本地地址
     *
     * @param filePath 导出文件路径（包含后缀）
     * @throws IOException IO异常
     */
    public void export(String filePath) throws IOException {
        switch (exportType) {
            case WORKBOOK:
                try (FileOutputStream fos = new FileOutputStream(filePath)) {
                    this.workbook.write(fos);
                }
                break;
            case CSV:
                // TODO csv export
                break;
        }
    }

    /**
     * 导出Excel数据到输出流
     *
     * @param os 输出流
     * @throws IOException IO异常
     */
    public void export(OutputStream os) throws IOException {
        switch (exportType) {
            case WORKBOOK:
                this.workbook.write(os);
                break;
            case CSV:
                // TODO csv export
                break;
        }
    }

    public ExportType getExportType() {
        return exportType;
    }

    public void setExportType(ExportType exportType) {
        this.exportType = exportType;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }
}
