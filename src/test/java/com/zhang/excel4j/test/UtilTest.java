package com.zhang.excel4j.test;

import com.zhang.excel4j.ExportUtil;
import com.zhang.excel4j.common.FieldAccessType;
import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.handler.ColumnHandler;
import com.zhang.excel4j.model.User;
import org.junit.Assert;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * author : zhangpan
 * date : 2018/1/30 15:30
 */
public class UtilTest {

    @Test
    public void test() {
        String bigDecimal = "2.0E02";
        boolean flg = Pattern.matches("^-?[1-9](\\.\\d+)?(E[+-]?\\d+)$", bigDecimal);
        boolean flg2 = Pattern.compile("^-?\\d+(\\.\\d+)?(E-?\\d+)?$").matcher(bigDecimal).find();
        System.out.println(flg);
        System.out.println(flg2);
    }

    @Test
    public void testGetterOrSetter() {
        try {
            User user = new User();
            user.setId("1");
            user.setUsername("张三");
            Method method = ColumnHandler.getterOrSetter(User.class, "username", FieldAccessType.GETTER);
            Object obj = method.invoke(user);
            Assert.assertTrue(user.getUsername().equals(obj));
        } catch (IntrospectionException | IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testExport01() {
        List<User> users = new ArrayList<>();
        int num = 0;
        while (num < 100) {
            User user = new User();
            user.setId((num + 1) + "");
            user.setUsername("张三");
            user.setPassword("sdfasdfasd");
            users.add(user);

            num++;
        }
        String path = "D:/Download/用户01.xls";
        try {
            ExportUtil.getInstance().getExportModel(users, User.class, null, null, WorkbookType.getWorkbookType("xls")).export(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testExport02() {
        List<Map<String, Object>> users = new ArrayList<>();
        int num = 0;
        while (num < 100) {
            Map<String, Object> user = new HashMap<>();
            user.put("id", (num + 1) + "");
            user.put("username", "张三");
            user.put("password", "sdfasdfasd");
            users.add(user);
            num++;
        }
        String path = "D:/Download/用户02.xls";
        try {
            ExportUtil.getInstance().getExportModel(users, User.class, null, null, WorkbookType.getWorkbookType("xls")).export(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    public void testExport03() throws Exception {
        List<User> users = new ArrayList<>();
        int num = 0;
        while (num < 100) {
            User user = new User();
            if (num % 2 == 0) {
                user.setId("D" + (num + 1));
                user.setUsername("张三");
                user.setPassword("dfd23456789");
            } else {
                user.setId("S" + (num + 1));
                user.setUsername("李四");
                user.setPassword("sdf33234555");
            }
            users.add(user);
            num++;
        }
        String path = "D:/Download/用户03.xlsx";
        File templateFile = getPropertiesFile("user_template01.xlsx");
        InputStream ins = new FileInputStream(templateFile);
        ExportUtil.getInstance().getExcelTemplate(ins, 0, users, null, User.class, null, false, "用户列表", WorkbookType.getWorkbookType("xlsx")).export(path);
    }

    @Test
    public void testExport04() throws Exception {
        List<User> users = new ArrayList<>();
        int num = 0;
        while (num < 100) {
            User user = new User();
            if (num % 2 == 0) {
                user.setId("D" + (num + 1));
                user.setUsername("张三");
                user.setPassword("dfd234567d89");
            } else {
                user.setId("S" + (num + 1));
                user.setUsername("李四");
                user.setPassword("sdf33234d555");
            }
            users.add(user);
            num++;
        }
        String path = "D:/Download/用户04.xlsx";
        File templateFile = getPropertiesFile("user_template02.xlsx");
        InputStream ins = new FileInputStream(templateFile);
        Map<String, Object> extendData = new HashMap<>();
        extendData.put("title", "用户列表");
        extendData.put("info", "说明");
        // ExportUtil.getInstance().exportList2Excel(path, ins, 0, users, extendData, User.class, null, false, "用户列表", false);
    }

    public File getPropertiesFile(String fileName) throws Exception {
        String dirPath = Thread.currentThread().getContextClassLoader().getResource("").toURI().getPath();
        return new File(dirPath, fileName);
    }
}
