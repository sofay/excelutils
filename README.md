# excelutils
for excel utils
## use like this
`model` like
```
 @ExcelExportInfo(fileNamePrefix = "fay_test",
            defaultColumnNameCellStyle = @CellStyle(font = @Font(fontName = "圆体"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22),
            defaultColumnValueCellStyle = @CellStyle(font = @Font(fontName = "娃娃体-简")),
            defaultLastRowCellStyle = @CellStyle(font = @Font(fontName = "魏碑-繁"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22))
     class ExportObj {
        @ExcelExportField(columnName = "名字啊", columnValueCellStyle = @CellStyle(fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22))
        private String name;

        @ExcelExportField(columnName = "年龄", order = 2, lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact")))
        private Integer age;

        @ExcelExportField(columnName = "简介", order = 3,
                columnNameCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "行楷-简")),
                columnValueCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "隶变-简")))
        private Object desc;
        }
        ......
```
then you can do
```
public static void main(String[] args) throws IOException {
        ExcelExportUtil.excelWriter(
                new ArrayList() {
                    {
                        add(ExportObj.builder().name("中国").age(70).desc("hello 字体测试").build());
                        add(ExportObj.builder().name("日本").age(3).desc("中abcdABCD文").build());
                        add(ExportObj.builder().name("汇总").age(70).build());
                    }
                })
                .write(new FileOutputStream(System.getProperty("user.home") + "/test.xls"));
    }
```
the export excel will like that
![avatar](http://sowcar.com/t6/687/1553159598x2890175145.png)
