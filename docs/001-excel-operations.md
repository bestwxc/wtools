# Excel Operations

``` java
### 从1.0.2版本开始支持
ExcelUtils.xssf()
    .sheet("sheet01")
        .row(0).fillRow(0, new String[]{"序号", "姓名", "年龄"}).sheet()
        .row(1).fillRow(0, new String[]{"1", "张三", "36"})
        .row(2)
            .cell(0).setCellValue("2")
            .cell(1).setCellValue("李四")
            .cell(2).setCellValue("27")
.write("abc.xlsx");
```