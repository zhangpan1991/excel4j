package com.zhang.excel4j.converter;

/**
 * 转换器接口
 *
 * author : zhangpan
 * date : 2018/1/29 11:39
 */
public interface Converter {

    /**
     * 读取Excel列内容转换
     *
     * @param object 待转换数据
     * @return 转换完成的结果
     */
    Object execRead(String object);

    /**
     * 写入Excel列内容转换
     *
     * @param object 待转换数据
     * @return  转换完成的结果
     */
    Object execWrite(Object object);
}
