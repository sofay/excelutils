package cn.fay.excel.annotations;

import cn.fay.excel.handle.DefaultExcelFieldTrans;

import java.lang.annotation.ElementType;
import java.lang.annotation.Repeatable;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午3:48.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(ExcelExportFields.class)
public @interface ExcelExportField {
    /**
     * 对应导出的excel中的列名 e.g. 姓名
     */
    String columnName();

    /**
     * 排序
     * 定义成double防止产品随意变更顺序
     * 对应导出excel中列的位置
     */
    double order() default 1;

    String usedForSerialNo() default ExcelExportInfo.DEFAULT_SERIAL_NO;

    /**
     * 使用该注解的SHEET
     */
    int[] usedForSheetIndex() default 0;

    /**
     * 默认值 (case sensitive)
     * support placeholder eg:
     * ${row}       : row index
     * ${column}    : column index
     *
     * support `` eg:
     * `1 + 1` => 2
     * `${column} - 1` : warn if expr contain ${column} the expr result will trans to ABCDEFG... 0 => A, 1 => B ...
     *                  eg: current column is 2(C) then ${column} => C and `${column} - 1` => 1 => B
     *
     */
    String defaultValue() default "";

    /**
     * 特殊字段的处理，比如类型等在导出时要处理成比较好理解的意思
     */
    String transClassName() default "cn.fay.excel.handle.DefaultExcelFieldTrans";

    /**
     * 同 {@link #transClassName()}
     */
    Class transClass() default DefaultExcelFieldTrans.class;

    /**
     * 列名单元格格式
     */
    CellStyle columnNameCellStyle() default @CellStyle;

    /**
     * 列值单元格格式
     */
    CellStyle columnValueCellStyle() default @CellStyle;

    /**
     * 当前列最后一行单元格格式
     */
    CellStyle lastRowCellStyle() default @CellStyle;

    /**
     * {@see defaultValue()}
     */
    String lastRowValue() default "";
}
