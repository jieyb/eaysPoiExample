package com.example.handler;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import com.example.entity.MyTest;
import org.springframework.stereotype.Component;

import java.util.StringJoiner;
import java.util.regex.Pattern;

/**
 * 自定义数据校验
 * 也可以将该类添加到spring容器管理
 */
//@Component
public class MyIExcelVerifyHandler implements IExcelVerifyHandler<MyTest> {
    /**
     * 校验方法
     * @param obj
     *            当前对象
     * @return
     */
    @Override
    public ExcelVerifyHandlerResult verifyHandler(MyTest obj) {
        // 用于添加错误提示
        StringJoiner joiner = new StringJoiner(",");
        // 校验名称不能为空
        if (obj.getName() == null || "".equals(obj.getName())) {
            joiner.add("姓名为空");
        }

        // 校验手机号不能为空或手机号格式等
        // 校验是为为11为数字
        String pattern = "^\\d{11}$";
        if (obj.getMobile() == null || "".equals(obj.getMobile()) || !Pattern.matches(pattern, obj.getMobile())) {
            joiner.add("手机号填写错误");
        }

        // 返回校验失败
        if (joiner.length() > 0) {
            return new ExcelVerifyHandlerResult(Boolean.FALSE, joiner.toString());
        }
        // 返回成功
        return new ExcelVerifyHandlerResult(Boolean.TRUE);
    }
}
