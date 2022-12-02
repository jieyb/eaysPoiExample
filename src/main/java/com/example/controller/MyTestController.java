package com.example.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.example.entity.MyTest;
import com.example.handler.MyIExcelVerifyHandler;
import com.example.service.MyTestService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

@RestController
@Slf4j
public class MyTestController {

    @Autowired
    private MyTestService myTestService;

    /**
     * excel 导入
     * @param file
     */
    @PostMapping("/test/import")
    public Map<String, Object> importExcel(MultipartFile file) {

        Map<String, Object> map = new HashMap<>();
        // 建立导入参数
        ImportParams params = new ImportParams();
        // 自定义参数处理
        // MyIExcelDataHandler dataHandler = new MyIExcelDataHandler();
        // params.setDataHandler(dataHandler);

        // 自定义参数校验
        params.setVerifyHandler(new MyIExcelVerifyHandler());
        try {
            // 导入解析
            ExcelImportResult<MyTest> result = ExcelImportUtil.importExcelMore(file.getInputStream(), MyTest.class, params);
            List<MyTest> list = result.getList();
            List<MyTest> failList = result.getFailList();
            StringJoiner joiner = new StringJoiner(";");
            for (MyTest t : failList) {
                Integer rowNum = t.getRowNum();
                String errorMsg = t.getErrorMsg();
                joiner.add("第 " + rowNum + " 行数据有误：" + errorMsg);
            }
            map.put("data", list);
            map.put("code", "2000");
            map.put("failList", failList);
            map.put("errorMsg", joiner.toString());
        } catch (Exception e) {
            log.error(e.getMessage());
        }
        return map;
    }

    /**
     * 导出 Excel
     */
    @GetMapping("/test/export")
    public void exportExcel(HttpServletResponse response) {
        // 查询数据
        List<MyTest> dataList = myTestService.getDataList();
        // 建立导出参数
        ExportParams params = new ExportParams();
        params.setSheetName("测试");
        try {
            Workbook workbook = ExcelExportUtil.exportExcel(params, MyTest.class, dataList);
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode("测试导出","UTF-8") + ".xls");
            ServletOutputStream outputStream = response.getOutputStream();
            workbook.write(outputStream);
        } catch (Exception e) {
            log.error(e.getMessage());
        }
    }

}
