package com.annotation.excel.utils;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

/**
 * 文件下载工具类
 *
 * @author : tdl
 * @date : 2018/12/5 下午1:46
 **/
public class DownloadUtil {
    /**
     * 下载模板
     *
     * @param response        响应
     * @param fileInputStream 文件输入流
     * @param fileName        文件名
     */
    public static void downloadFile(HttpServletResponse response, InputStream fileInputStream, String fileName) throws IOException {
        OutputStream os = response.getOutputStream();
        InputStream inputStream = null;
        BufferedInputStream bufferedInputStream = null;
        BufferedOutputStream bouts = null;

        byte[] buffer = new byte[1024];

        try {
            inputStream = fileInputStream;
            // 设置强制下载不打开
            response.setContentType("application/force-download");
            // 设置文件名
            response.addHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName, "UTF-8"));

            bufferedInputStream = new BufferedInputStream(inputStream);
            bouts = new BufferedOutputStream(os);
            int i = bufferedInputStream.read(buffer);
            while (i != -1) {
                bouts.write(buffer, 0, i);
                i = bufferedInputStream.read(buffer);
            }
            // 这里一定要调用flush()方法
            bouts.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (bufferedInputStream != null) {
                bufferedInputStream.close();
                bufferedInputStream = null;
            }

            if (inputStream != null) {
                inputStream.close();
                inputStream = null;
            }
        }
    }
}
