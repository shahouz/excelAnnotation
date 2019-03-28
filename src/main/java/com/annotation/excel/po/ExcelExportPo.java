package com.annotation.excel.po;

import lombok.Data;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * excel导出入参
 *
 * @author : tdl
 * @date : 2019/3/28 下午5:24
 **/
@Data
public class ExcelExportPo {
    private HttpServletRequest request;

    private HttpServletResponse response;

    private String excelName;

    private List<Object> list;
}
