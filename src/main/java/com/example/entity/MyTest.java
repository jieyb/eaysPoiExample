package com.example.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;
import lombok.Data;

/**
 * 继承 IExcelModel 方便获取数据校验返回的错误信息
 * 继承 IExcelDataModel 方便获取行数
 */
@ExcelTarget("myTest")
@Data
public class MyTest implements IExcelModel, IExcelDataModel {

    /**
     * excel 注解中的name对应excel表格的表头
     */
    @Excel(name = "姓名")
    private String name;

    /**
     * 数据转换，导入时填写男，则转成1；导出则相反
     */
    @Excel(name = "性别", replace = {"男_1", "女_2"}, isImportField = "true")
    private Integer sex;

    @Excel(name = "年龄")
    private Integer age;

    @Excel(name = "手机号")
    private String mobile;

    /**
     * 地市名称
     */
    @Excel(name = "所在地市")
    private String cityName;

    /**
     * 地市编码
     */
    private String cityCode;

    /**
     * 行数
     */
    private Integer rowNum;

    /**
     * 错误提示信息
     */
    private String errorMsg;

    @Override
    public Integer getRowNum() {
        return rowNum;
    }

    @Override
    public void setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
    }

    @Override
    public String getErrorMsg() {
        return errorMsg;
    }

    @Override
    public void setErrorMsg(String errorMsg) {
        this.errorMsg = errorMsg;
    }
}
