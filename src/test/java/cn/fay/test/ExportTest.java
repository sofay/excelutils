package cn.fay.test;

import cn.fay.excel.ExcelExportUtil;
import cn.fay.excel.annotations.CellStyle;
import cn.fay.excel.annotations.ExcelExportField;
import cn.fay.excel.annotations.ExcelExportInfo;
import cn.fay.excel.annotations.Font;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 16:38.
 */
@SuppressWarnings("all")
public class ExportTest {

    @ExcelExportInfo(fileNamePrefix = "fay_test",
            defaultColumnNameCellStyle = @CellStyle(font = @Font(fontName = "圆体"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22),
            defaultColumnValueCellStyle = @CellStyle(font = @Font(fontName = "娃娃体-简")),
            defaultLastRowCellStyle = @CellStyle(font = @Font(fontName = "魏碑-繁"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22))
    static class ExportObj {
        @ExcelExportField(columnName = "名字啊", columnValueCellStyle = @CellStyle(fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22))
        private String name;

        @ExcelExportField(columnName = "年龄", order = 2, lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact")))
        private Integer age;

        @ExcelExportField(columnName = "简介", order = 3,
                columnNameCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "行楷-简")),
                columnValueCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "隶变-简")))
        private Object desc;

        public static Builder builder() {
            return new Builder();
        }

        public static class Builder {
            private ExportObj obj;

            private Builder() {
                this.obj = new ExportObj();
            }

            public Builder name(String name) {
                obj.name = name;
                return this;
            }

            public Builder age(int age) {
                obj.age = age;
                return this;
            }

            public Builder desc(Object desc) {
                obj.desc = desc;
                return this;
            }

            public ExportObj build() {
                return obj;
            }
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }

        public Object getDesc() {
            return desc;
        }

        public void setDesc(Object desc) {
            this.desc = desc;
        }
    }

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
        System.out.println("done");
    }
}
