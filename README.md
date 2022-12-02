# Spring Boot整合EaysPoi实现基本导入导出

## 1、引入依赖
``` xml
<dependencies>
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter</artifactId>
        <version>2.3.7.RELEASE</version>
    </dependency>
    
    <dependency>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-web</artifactId>
        <version>2.3.7.RELEASE</version>
    </dependency>
    <dependency>
        <groupId>javax.servlet</groupId>
        <artifactId>servlet-api</artifactId>
        <version>2.5</version>
    </dependency>
    
    <dependency>
        <groupId>org.projectlombok</groupId>
        <artifactId>lombok</artifactId>
        <version>1.16.10</version>
    </dependency>
    
    <!-- easyPoi依赖 -->
    <dependency>
        <groupId>cn.afterturn</groupId>
        <artifactId>easypoi-spring-boot-starter</artifactId>
        <version>4.3.0</version>
    </dependency>
    <!-- 建议只用start -->
    <!--<dependency>
        <groupId>cn.afterturn</groupId>
        <artifactId>easypoi-base</artifactId>
        <version>4.3.0</version>
    </dependency>
    <dependency>
        <groupId>cn.afterturn</groupId>
        <artifactId>easypoi-web</artifactId>
        <version>4.3.0</version>
    </dependency>
    <dependency>
        <groupId>cn.afterturn</groupId>
        <artifactId>easypoi-annotation</artifactId>
        <version>4.3.0</version>
    </dependency>-->
</dependencies>
```
**新建entity数据解析**
``` java
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
     * replace: 数据转换，导入时填写男，则转成1；导出则相反
     */
    @Excel(name = "性别", replace = {"男_1", "女_2"}, isImportField = "true")
    private Integer sex;

    @Excel(name = "年龄")
    private Integer age;

    @Excel(name = "手机号")
    private String mobile;

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
```
## 2、导入
### 2.1 基础导入
**实现**
``` java
public static void importExcel(InputStream inputStream) {
    // 导入参数
    // 该类可设置 表格标题行数,默认0、表头行数,默认1、开始读取的sheet位置,默认为0、上传表格需要读取的sheet 数量,默认为1 等......
    // 以上参数设置均可进入 ImportParams 源码查看相应注释，这里使用默认；
    ImportParams importParams = new ImportParams();
    try {
        // 导入数据解析
        ExcelImportResult<MyTest> importResult = ExcelImportUtil.importExcelMore(inputStream, MyTest.class, importParams);

        // 解析成功的数据集
        List<MyTest> successList = importResult.getList();
        log.info("成功的数据集：{}", Arrays.toString(successList.toArray()));
        // 解析失败的数据集
        List<MyTest> failList = importResult.getFailList();
        log.info("失败的数据集：{}", Arrays.toString(failList.toArray()));
        for (MyTest myTest : failList) {
            Integer rowNum = myTest.getRowNum();
            String errorMsg = myTest.getErrorMsg();
            log.info("第 {} 行数据填写错误：{}", rowNum, errorMsg);
        }
        // TODO 入库操作

    } catch (Exception e) {
        e.printStackTrace();
    }
}
```
### 2.2 自定义数据校验
**新建自定义数据校验类继承`IExcelVerifyHandler`**
``` java
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
```
**导入的`ImportParams`参数设置自定义校验类**
``` java
// 设置导入数据自定义校验
importParams.setVerifyHandler(new MyIExcelVerifyHandler());
```
效果
```text
16:10:42.065 [main] INFO com.example.controller.ExcelTest - 成功的数据集：[MyTest(name=小红, sex=2, age=22, mobile=15778293841, rowNum=1, errorMsg=null)]
16:10:42.065 [main] INFO com.example.controller.ExcelTest - 失败的数据集：[MyTest(name=null, sex=1, age=23, mobile=123456789, rowNum=2, errorMsg=姓名为空,手机号填写错误), MyTest(name=小黑, sex=1, age=23, mobile=null, rowNum=3, errorMsg=手机号填写错误)]
16:10:42.065 [main] INFO com.example.controller.ExcelTest - 第 2 行数据填写错误：姓名为空,手机号填写错误
16:10:42.065 [main] INFO com.example.controller.ExcelTest - 第 3 行数据填写错误：手机号填写错误
```
### 2.3 数据自定义处理
**新建类继承`IExcelDataHandler`**
``` java
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
```
**完善`ImportParams`参数设置**
``` java
// 设置导入数据自定义处理
MyIExcelDataHandler dataHandler = new MyIExcelDataHandler();
// 测试使用造数
Map<String, String> cityMap = new HashMap<>();
cityMap.put("南宁市", "450100");
cityMap.put("崇左市", "451400");
dataHandler.setCityMap(cityMap);
// 设置需要处理的字段，自定义那里只会处理这里传入的字段
dataHandler.setNeedHandlerFields(new String[]{"所在地市"});
importParams.setDataHandler(dataHandler);
```

## 3、导出
``` java
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
        // 这里直接导出当前项目目录
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
```