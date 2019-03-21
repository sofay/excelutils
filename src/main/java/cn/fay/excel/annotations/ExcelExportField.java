package cn.fay.excel.annotations;

import cn.fay.excel.handle.DefaultExcelFieldTrans;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午3:48.
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
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

    /**
     * 默认值
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
     * 列名单元格类型
     */
    int columnNameCellType() default Cell.CELL_TYPE_STRING;

    /**
     * 列值单元格类型
     */
    int columnValueCellType() default Cell.CELL_TYPE_STRING;
}
