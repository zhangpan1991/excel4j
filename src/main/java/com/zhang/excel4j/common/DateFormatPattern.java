package com.zhang.excel4j.common;

/**
 * 日期格式化规范常量
 *
 * author : zhangpan
 * date : 2018/1/31 16:23
 */
public interface DateFormatPattern {

    String DATE_FORMAT_DAY = "yyyy-MM-dd";
    String DATE_FORMAT_DAY_2 = "yyyy/MM/dd";
    String TIME_FORMAT_SEC = "HH:mm:ss";
    String DATE_FORMAT_SEC = "yyyy-MM-dd HH:mm:ss";
    String DATE_FORMAT_MSEC = "yyyy-MM-dd HH:mm:ss.SSS";
    String DATE_FORMAT_MSEC_T = "yyyy-MM-dd'T'HH:mm:ss.SSS";
    String DATE_FORMAT_MSEC_T_Z = "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'";
    String DATE_FORMAT_DAY_SIMPLE = "y/M/d";
}
