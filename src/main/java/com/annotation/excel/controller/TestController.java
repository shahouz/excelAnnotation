package com.annotation.excel.controller;

import com.annotation.excel.po.ExcelExportPo;
import com.annotation.excel.utils.ExcelUtil;
import com.annotation.excel.vo.ExcelExportVo;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

/**
 * 数据中心
 *
 * @author : tdl
 * @date : 2019/3/25 下午6:42
 **/
@Slf4j
@RestController
@RequestMapping("test")
public class TestController {

    @GetMapping("download/excel")
    public void downloadOperatorData(HttpServletRequest request, HttpServletResponse response) {
        ExcelExportVo vo1 = new ExcelExportVo();
        vo1.setName("黑黑驴");
        vo1.setDepartment("研发部");
        vo1.setCellphone("13606031000");
        ExcelExportVo vo2 = new ExcelExportVo();
        vo2.setName("灰灰马");
        vo2.setDepartment("财务部");
        vo2.setCellphone("13606031001");

        List<Object> list = new ArrayList<>();
        list.add(vo1);
        list.add(vo2);

        ExcelExportPo params = new ExcelExportPo();
        params.setExcelName("运营人员数据");
        params.setList(list);
        params.setRequest(request);
        params.setResponse(response);

        try {
            ExcelUtil.exportExcel(params);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
