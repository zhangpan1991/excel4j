package com.zhang.excel4j.model;

import com.zhang.excel4j.annotation.Column;
import com.zhang.excel4j.annotation.GroupBy;
import com.zhang.excel4j.common.GroupType;

/**
 * author : zhangpan
 * date : 2018/1/26 14:38
 */
public class User {

    @Column(value = "编号", order = 1)
    private String id;

    @Column(value = "用户名", groupType = GroupType.MUST, order = 2)
    @GroupBy({"login","info"})
    private String username;

    @Column(value = "密码", order = 3)
    @GroupBy("login")
    private String password;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }
}
