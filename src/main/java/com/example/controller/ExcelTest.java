package com.example.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import com.example.entity.MyTest;
import com.example.handler.MyIExcelDataHandler;
import com.example.handler.MyIExcelVerifyHandler;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.core.io.ClassPathResource;

import java.io.*;
import java.util.*;

@Slf4j
public class ExcelTest {

    public static void main(String[] args) {
        try {
            // 测试导入
            ClassPathResource classPathResource = new ClassPathResource("excelFile/MyTest导入模板.xlsx");
            InputStream inputStream = classPathResource.getInputStream();
            importExcel(inputStream);

            // 测试导出
            // exportExcel();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 导入
     * @param inputStream
     */
    public static void importExcel(InputStream inputStream) {
        // 导入参数
        // 该类可设置 表格标题行数,默认0、表头行数,默认1、开始读取的sheet位置,默认为0、上传表格需要读取的sheet 数量,默认为1 等......
        // 以上参数设置均可进入 ImportParams 源码查看相应注释，这里使用默认；
        ImportParams importParams = new ImportParams();
        try {
            // *** 数据处理会在数据校验之前执行
            // 设置导入数据自定义处理
            MyIExcelDataHandler dataHandler = new MyIExcelDataHandler();
            // 测试使用造数
            Map<String, String> cityMap = new HashMap<>();
            cityMap.put("南宁市", "450100");
            cityMap.put("崇左市", "451400");
            dataHandler.setCityMap(cityMap);
            // 设置需要处理的字段
            dataHandler.setNeedHandlerFields(new String[]{"所在地市"});
            importParams.setDataHandler(dataHandler);
            // 设置导入数据自定义校验
            importParams.setVerifyHandler(new MyIExcelVerifyHandler());
            // 导入数据解析
            ExcelImportResult<MyTest> importResult = ExcelImportUtil.importExcelMore(inputStream, MyTest.class, importParams);

            // 拿到数据后可以自己做其他操作，如：入库 等
            // 解析成功的数据集
            List<MyTest> successList = importResult.getList();
            log.info("成功的数据集：{}", Arrays.toString(successList.toArray()));
            // 解析失败的数据集，可用于回显前端展示，提示用户那里填错了需要重新填写等
            List<MyTest> failList = importResult.getFailList();
            log.info("失败的数据集：{}", Arrays.toString(failList.toArray()));
            for (MyTest myTest : failList) {
                Integer rowNum = myTest.getRowNum();
                String errorMsg = myTest.getErrorMsg();
                log.info("第 {} 行数据填写错误：{}", rowNum, errorMsg);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 导出
     */
    public static void exportExcel() throws IOException {
        // 导出参数
        // 具体属性设置参考源码注释
        ExportParams exportParams = new ExportParams();
        // 设置表单名称
        exportParams.setSheetName("测试");
        // 导出版本，默认 xls
        exportParams.setType(ExcelType.XSSF);
        // 获取数据
        List<MyTest> list = getDataList();

        FileOutputStream fileOutputStream = null;
        try {
            // 导出
            // 大数据量读取可参考官方文档编写，这里只做基础使用演示
            Workbook workbook = ExcelExportUtil.exportExcel(exportParams, MyTest.class, list);
            // 直接导出当前项目目录
            // 默认导出版本 xls，可通过 ExportParams 参数设置
            fileOutputStream = new FileOutputStream("测试导出.xlsx");
            workbook.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fileOutputStream != null) {
                fileOutputStream.close();
            }
        }
    }

    /**
     * 模拟数据
     * @return
     */
    public static List<MyTest> getDataList() {
        List<MyTest> list = new ArrayList<>();
        MyTest myTest = new MyTest();
        myTest.setName("小红");
        myTest.setSex(1);
        myTest.setAge(22);
        myTest.setMobile("15778192883");

        MyTest myTest2 = new MyTest();
        myTest2.setName("小黑");
        myTest2.setSex(1);
        myTest2.setAge(22);
        myTest2.setMobile("15778192882");
        list.add(myTest2);
        return list;
    }

}
