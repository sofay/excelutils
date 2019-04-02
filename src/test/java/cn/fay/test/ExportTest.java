package cn.fay.test;

import cn.fay.excel.ExcelExportExecutor;
import cn.fay.excel.annotations.CellStyle;
import cn.fay.excel.annotations.ExcelExportField;
import cn.fay.excel.annotations.ExcelExportInfo;
import cn.fay.excel.annotations.Font;
import lombok.Builder;
import lombok.Data;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 16:38.
 */
@SuppressWarnings("all")
public class ExportTest {

    @ExcelExportInfo(
            defaultColumnNameCellStyle = @CellStyle(font = @Font(fontName = "圆体"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22),
            defaultColumnValueCellStyle = @CellStyle(font = @Font(fontName = "娃娃体-简")),
            defaultLastRowCellStyle = @CellStyle(font = @Font(fontName = "魏碑-繁"), fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22))
    @Data
    @Builder
    static class ExportObj {
        @ExcelExportField(columnName = "名字啊", columnValueCellStyle = @CellStyle(fillPattern = HSSFCellStyle.SOLID_FOREGROUND, fillForegroundColor = 22), lastRowValue = "汇总")
        private String name;

        @ExcelExportField(columnName = "年龄", order = 2,
                columnValueCellStyle = @CellStyle(cellType = Cell.CELL_TYPE_NUMERIC, alignment = HSSFCellStyle.ALIGN_RIGHT),
                lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact"), cellType = Cell.CELL_TYPE_FORMULA),
                lastRowValue = "faySum()")
        private Integer age;

        @ExcelExportField(columnName = "年龄2", order = 2.1,
                columnValueCellStyle = @CellStyle(cellType = Cell.CELL_TYPE_NUMERIC),
                lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact"), cellType = Cell.CELL_TYPE_FORMULA),
                lastRowValue = "faySum()")
        private Integer age2;

        @ExcelExportField(columnName = "测试列(年龄/年龄2)", order = 2.5,
                columnValueCellStyle = @CellStyle(cellType = Cell.CELL_TYPE_FORMULA),
                defaultValue = "magic1()",
                lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact"), cellType = Cell.CELL_TYPE_FORMULA),
                lastRowValue = "magic1()")
        private Object test;

        @ExcelExportField(columnName = "测试列2(当前年龄/最大年龄)", order = 2.6,
                columnValueCellStyle = @CellStyle(cellType = Cell.CELL_TYPE_FORMULA),
                defaultValue = "magic2()",
                lastRowCellStyle = @CellStyle(font = @Font(fontName = "Impact")),
                lastRowValue = "100%")
        private Object test2;

        @ExcelExportField(columnName = "简介", order = 3,
                columnNameCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "行楷-简")),
                columnValueCellStyle = @CellStyle(alignment = HSSFCellStyle.ALIGN_RIGHT, font = @Font(fontName = "隶变-简")))
        private Object desc;
    }

    public static void main(String[] args) throws IOException {
        Logger.getRootLogger().setLevel(Level.DEBUG);
        ExcelExportExecutor.excelWriter(
                new ArrayList() {
                    {
                        add(ExportObj.builder().name("中国").age(70).age2(99).desc("hello 字体测试").build());
                        add(ExportObj.builder().name("日本").age(3).age2(1).desc("中abcdABCD文").build());
                    }
                })
                .write(new FileOutputStream(System.getProperty("user.home") + "/test2.xls"));
        System.out.println("done");
    }
}
