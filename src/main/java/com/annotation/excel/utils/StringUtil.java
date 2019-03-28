package com.annotation.excel.utils;

/**
 * 字符串处理工具
 *
 * @author : tdl
 * @date : 2019/3/28 下午6:22
 **/
public class StringUtil {
    public static boolean isBlank(String... args) {
        for (String arg : args) {
            if (arg == null) {
                continue;
            }
            if (arg.trim() == "" || "".equals(arg.trim())) {
                continue;
            }
            if ("null".equalsIgnoreCase(arg.trim())) {
                continue;
            }

            return false;
        }

        return true;
    }

    /**
     * Object对象转换为字符串，如果对象为空，则返回空串，否则返回该对象转换的字符串
     *
     * @param object 被转换对象
     * @return 转换后的字符串
     */
    public static String getObjectToString(Object object) {
        String str = "";
        try {
            if (object != null && !isBlank(object.toString())) {
                str = object.toString();
            }
        } catch (Exception e) {
            System.out.println("Object转换为字符串出错!");
            e.printStackTrace();
        }
        return str;
    }
}
