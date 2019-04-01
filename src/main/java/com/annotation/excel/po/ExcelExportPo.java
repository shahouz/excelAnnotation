package com.annotation.excel.po;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * excel导出入参
 *
 * @author : tdl
 * @date : 2019/3/28 下午5:24
 **/
public class ExcelExportPo {
    private HttpServletRequest request;

    private HttpServletResponse response;

    private String excelName;

    private List<Object> list;

    /**
     * 动态控制要显示的字段
     */
    private List<String> columnFilter;

    public HttpServletRequest getRequest() {
        return request;
    }

    public void setRequest(HttpServletRequest request) {
        this.request = request;
    }

    public HttpServletResponse getResponse() {
        return response;
    }

    public void setResponse(HttpServletResponse response) {
        this.response = response;
    }

    public String getExcelName() {
        return excelName;
    }

    public void setExcelName(String excelName) {
        this.excelName = excelName;
    }

    public List<Object> getList() {
        return list;
    }

    public void setList(List<Object> list) {
        this.list = list;
    }

    public List<String> getColumnFilter() {
        return columnFilter;
    }

    public void setColumnFilter(List<String> columnFilter) {
        this.columnFilter = columnFilter;
    }
}
