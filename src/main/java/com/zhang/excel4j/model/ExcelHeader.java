package com.zhang.excel4j.model;

import com.zhang.excel4j.converter.Converter;

/**
 * 用来存储文件标题的对象，通过该对象可以获取标题和方法的对应关系
 *
 * author : zhangpan
 * date : 2018/1/25 16:49
 */
public class ExcelHeader implements Comparable<ExcelHeader> {

    /**
     * 标题名称
     */
    private String title;

    /**
     * 标题的排序顺序
     */
    private Double order;

    /**
     * 数据转换器
     */
    private Converter converter;

    /**
     * 注解域
     */
    private String filed;

    /**
     * 属性类型
     */
    private Class<?> filedClazz;

    public ExcelHeader() {
    }

    public ExcelHeader(String title, Double order, Converter converter, String filed, Class<?> filedClazz) {
        this.title = title;
        this.order = order;
        this.converter = converter;
        this.filed = filed;
        this.filedClazz = filedClazz;
    }

    @Override
    public int compareTo(ExcelHeader o) {
        return this.order.compareTo(o.order);
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Double getOrder() {
        return order;
    }

    public void setOrder(Double order) {
        this.order = order;
    }

    public Converter getConverter() {
        return converter;
    }

    public void setConverter(Converter converter) {
        this.converter = converter;
    }

    public String getFiled() {
        return filed;
    }

    public void setFiled(String filed) {
        this.filed = filed;
    }

    public Class<?> getFiledClazz() {
        return filedClazz;
    }

    public void setFiledClazz(Class<?> filedClazz) {
        this.filedClazz = filedClazz;
    }
}
