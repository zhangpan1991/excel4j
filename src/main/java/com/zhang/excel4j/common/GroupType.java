package com.zhang.excel4j.common;

/**
 * 分组类型枚举类
 * {@link GroupType#DEFAULT} 默认分组类型，没有分组时，按任意分组匹配，有分组时，按已有分组匹配
 * {@link GroupType#NON} 不分组，但匹配任意分组
 * {@link GroupType#ALWAYS} 所有分组都可匹配，有分组按分组的排序值
 * {@link GroupType#MUST} 必须分组，不分组无法匹配
 *
 * author : zhangpan
 * date : 2018/1/26 18:16
 */
public enum GroupType {
    DEFAULT, NON, ALWAYS, MUST
}
