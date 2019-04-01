package com.annotation.excel.constant;

/**
 * 文件常量
 *
 * @author : tdl
 * @date : 2018/12/3 上午10:23
 **/
public class FileConstants {
    /**
     * 可上传文件格式
     */
    public final static String[] LEGAL_FORMATE = {".rar", ".zip", ".doc", ".docx", ".xls", ".xlsx", ".jpg", ".jpeg", ".png", ".txt"};

    /**
     * 单次上传最大文件尺寸(字节)
     */
    public final static Integer MAX_SIZE = 20971520;

    /**
     * 导出数据文件格式
     */
    public final static String EXPORT_FORMATE = ".xlsx";

    /**
     * 单次下载最大数量
     */
    public static final Integer DOWNLOAD_MAX = 5000;

    public final static String EXPORT_CACHE_FILE_PATH = "excel";
}
