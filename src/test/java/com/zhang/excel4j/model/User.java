package com.zhang.excel4j.model;

import com.zhang.excel4j.annotation.Column;
import com.zhang.excel4j.annotation.GroupBy;
import com.zhang.excel4j.common.GroupType;

import java.util.Date;

/**
 * author : zhangpan
 * date : 2018/1/26 14:38
 */
public class User {

    @Column(value = "序号", order = 0)
    private Integer num;

    @Column(value = "编号", order = 1)
    private String id;

    @Column(value = "用户名", groupType = GroupType.MUST, order = 2)
    @GroupBy({"login","info"})
    private String uesrname;

    @Column(value = "密码", order = 3)
    @GroupBy("login")
    private String password;

    private String realname;

    private String department;

    private String telphone;

    private Date birthday;

    public Integer getNum() {
        return num;
    }

    public void setNum(Integer num) {
        this.num = num;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getUesrname() {
        return uesrname;
    }

    public void setUesrname(String uesrname) {
        this.uesrname = uesrname;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public String getRealname() {
        return realname;
    }

    public void setRealname(String realname) {
        this.realname = realname;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getTelphone() {
        return telphone;
    }

    public void setTelphone(String telphone) {
        this.telphone = telphone;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }
}
