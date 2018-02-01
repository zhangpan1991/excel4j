package com.zhang.excel4j.util;

import com.zhang.excel4j.common.DateFormatPattern;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Pattern;

/**
 * 日期时间相关工具
 *
 * author : zhangpan
 * date : 2018/1/29 19:06
 */
public class DateUtil {

    /**
     * 匹配yyyy-MM-dd
     */
    private static final String DATE_REG = "^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])$";
    /**
     * 匹配yyyy/MM/dd
     */
    private static final String DATE_REG_2 = "^[1-9]\\d{3}/(0[1-9]|1[0-2])/(0[1-9]|[1-2][0-9]|3[0-1])$";
    /**
     * 匹配y/M/d
     */
    private static final String DATE_REG_SIMPLE_2 = "^[1-9]\\d{3}/([1-9]|1[0-2])/([1-9]|[1-2][0-9]|3[0-1])$";
    /**
     * 匹配HH:mm:ss
     */
    private static final String TIME_SEC_REG = "^(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d$";
    /**
     * 匹配yyyy-MM-dd HH:mm:ss
     */
    private static final String DATE_TIME_REG = "^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])\\s" +
            "(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d$";
    /**
     * 匹配yyyy-MM-dd HH:mm:ss.SSS
     */
    private static final String DATE_TIME_MSEC_REG = "^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])\\s" +
            "(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d\\.\\d{3}$";
    /**
     * 匹配yyyy-MM-dd'T'HH:mm:ss.SSS
     */
    private static final String DATE_TIME_MSEC_T_REG = "^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])T" +
            "(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d\\.\\d{3}$";
    /**
     * 匹配yyyy-MM-dd'T'HH:mm:ss.SSS'Z'
     */
    private static final String DATE_TIME_MSEC_T_Z_REG = "^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])T" +
            "(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d\\.\\d{3}Z$";

    /**
     * 将日期时间转换为指定格式的字符串
     *
     * @param date   日期时间
     * @param format 指定格式化类型
     * @return 返回格式化后的时间字符串
     */
    public static String date2Str(Date date, String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(date);
    }

    /**
     * 将日期时间转换为默认为[yyyy-MM-dd HH:mm:ss]类型的字符串
     *
     * @param date 日期时间
     * @return 返回格式化后的时间字符串
     */
    public static String date2Str(Date date) {
        SimpleDateFormat sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_SEC);
        return sdf.format(date);
    }

    /**
     * 根据给出的格式化类型将时间字符串转为类型
     *
     * @param strDate 时间字符串
     * @param format  格式化类型
     * @return 返回{@link java.util.Date}类型
     */
    public static Date str2Date(String strDate, String format) {
        Date date = null;
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        try {
            date = sdf.parse(strDate);
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return date;
    }

    /**
     * 字符串时间自动匹配转为日期时间，未找到匹配类型返回null
     *
     * @param strDate 时间字符串
     * @return Date 日期时间
     * @throws ParseException 异常
     */
    public static Date str2Date(String strDate) throws ParseException {
        strDate = strDate.trim();
        SimpleDateFormat sdf = null;
        if (Pattern.matches(strDate, DATE_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_DAY);
        }
        if (Pattern.matches(strDate, DATE_REG_2)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_DAY_2);
        }
        if (Pattern.matches(strDate, DATE_REG_SIMPLE_2)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_DAY_SIMPLE);
        }
        if (Pattern.matches(strDate, TIME_SEC_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.TIME_FORMAT_SEC);
        }
        if (Pattern.matches(strDate, DATE_TIME_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_SEC);
        }
        if (Pattern.matches(strDate, DATE_TIME_MSEC_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_MSEC);
        }
        if (Pattern.matches(strDate, DATE_TIME_MSEC_T_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_MSEC_T);
        }
        if (Pattern.matches(strDate, DATE_TIME_MSEC_T_Z_REG)) {
            sdf = new SimpleDateFormat(DateFormatPattern.DATE_FORMAT_MSEC_T_Z);
        }
        if (null != sdf) {
            return sdf.parse(strDate);
        }
        return null;
    }

}
