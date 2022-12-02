package com.example.handler;

import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import com.example.entity.MyTest;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.util.HashMap;
import java.util.Map;

/**
 * 自定义数据处理
 */
public class MyIExcelDataHandler implements IExcelDataHandler<MyTest> {

    /**
     * 地市数据
     */
    private Map<String, String> cityMap = new HashMap<>();

    /**
     * 处理的字段
     */
    private String[] needHandlerFields;

    /**
     * 导出的数据处理
     * @param obj
     *            当前对象
     * @param name
     *            当前字段名称
     * @param value
     *            当前值
     * @return
     */
    @Override
    public Object exportHandler(MyTest obj, String name, Object value) {
        return value;
    }

    @Override
    public String[] getNeedHandlerFields() {
        return needHandlerFields;
    }

    /**
     * 导入数据处理
     * @param obj
     *            当前对象
     * @param name
     *            当前字段名称
     * @param value
     *            当前值
     * @return
     */
    @Override
    public Object importHandler(MyTest obj, String name, Object value) {
        // 这里会默认按照表格的表头顺序进行，例如：模板表头为 姓名、性别、年龄，则依次进入的就是 姓名、性别、年龄
        // 这里的name 是标题名称：姓名、性别...
        // 这里的obj属性值并不是全部数据都有；例如：当前name等于姓名，那按顺序对应的性别、年龄 字段是还没有值的
        if ("所在地市".equals(name)) {
            // 获取填充地市的code
            obj.setCityCode(cityMap.get((String) value));
        }
        return value;
    }

    /**
     * 需要处理的字段名称
     * @param fields
     */
    @Override
    public void setNeedHandlerFields(String[] fields) {
        this.needHandlerFields = fields;
    }

    @Override
    public void setMapValue(Map<String, Object> map, String originKey, Object value) {
        map.put(originKey, value);
    }

    @Override
    public Hyperlink getHyperlink(CreationHelper creationHelper, MyTest obj, String name, Object value) {
        return null;
    }

    public MyIExcelDataHandler setCityMap(Map<String, String> cityMap) {
        if (cityMap != null) {
            this.cityMap = cityMap;
        }
        return this;
    }
}
