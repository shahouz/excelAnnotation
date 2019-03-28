package com.annotation.excel.vo;

import com.annotation.excel.annotation.Excel;
import lombok.Data;

/**
 * excel导出 - 出参结构
 *
 * @author : tdl
 * @date : 2019/3/28 下午6:11
 **/
@Data
public class ExcelExportVo {
    @Excel(title = "姓名", width = 2000)
    private String name;

    @Excel(title = "部门", width = 2000)
    private String department;

    @Excel(title = "手机", width = 4000)
    private String cellphone;
}
