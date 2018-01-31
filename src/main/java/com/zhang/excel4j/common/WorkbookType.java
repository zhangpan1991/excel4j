package com.zhang.excel4j.common;

/**
 * 工作簿类型
 *
 * author : zhangpan
 * date : 2018/1/30 11:39
 */
public enum WorkbookType {
    XSSF(false, "xlsx"), HSSF(false, "xls"), SXSSF(true, "xlsx");

    /**
     * 是否特殊类型
     */
    private boolean special;

    /**
     * 后缀
     */
    private String suffix;

    WorkbookType(boolean special, String suffix) {
        this.suffix = suffix;
    }

    /**
     * 通过后缀获取工作簿类型
     *
     * @param suffix 后缀
     * @return 工作簿类型
     */
    public static WorkbookType getWorkbookType(String suffix) {
        for (WorkbookType workbooType : WorkbookType.values()) {
            if (!workbooType.getSpecial() && workbooType.getSuffix().equals(suffix)) {
                return workbooType;
            }
        }
        return null;
    }

    public boolean getSpecial() {
        return special;
    }

    public void setSpecial(boolean special) {
        this.special = special;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }
}
