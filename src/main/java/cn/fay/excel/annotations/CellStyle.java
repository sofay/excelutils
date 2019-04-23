package cn.fay.excel.annotations;

import cn.fay.excel.handle.CellStyleHandler;
import cn.fay.excel.handle.UseCellStyleMethodHandler;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-20 14:36.
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface CellStyle {
    String STRING_DEFAULT_VALUE = "Arial";
    short SHORT_DEFAULT_VALUE = 0;
    short SHORT_DEFAULT_VALUE_1 = 1;
    short SHORT_DEFAULT_VALUE_10 = 10;

    Font font() default @Font;

    short fontHeightInPoints() default SHORT_DEFAULT_VALUE_10;

    short alignment() default SHORT_DEFAULT_VALUE_1;

    short borderBottom() default SHORT_DEFAULT_VALUE_1;

    short borderLeft() default SHORT_DEFAULT_VALUE_1;

    short borderRight() default SHORT_DEFAULT_VALUE_1;

    short borderTop() default SHORT_DEFAULT_VALUE_1;

    short fillPattern() default SHORT_DEFAULT_VALUE;

    short fillForegroundColor() default SHORT_DEFAULT_VALUE;

    short dataFormat() default SHORT_DEFAULT_VALUE;

    int cellType() default Cell.CELL_TYPE_STRING;

    String cellStyleHandlerClassName() default "com.raycloud.kmsy.domain.UseCellStyleMethodHandler";

    Class<? extends CellStyleHandler> cellStyleHandlerClass() default UseCellStyleMethodHandler.class;


    /**
     * 给出几种默认的单元格格式供选择
     */
    interface Value {
        @CellStyle()
        Object STRING_CELL_STYLE = null;
        @CellStyle()
        Object COMMON_CELL_STYLE = null;
        @CellStyle
        Object DEFAULT_CELL_STYLE = null;

        class Impl implements Value {
            private static CellStyle defaultCS = null;

            public static CellStyle defaultCellStyle() {
                if (defaultCS != null) {
                    return defaultCS;
                }
                return defaultCS = getCellStyleFromField("DEFAULT_CELL_STYLE");
            }

            private static CellStyle getCellStyleFromField(String fieldName) {
                try {
                    return Value.class.getField(fieldName).getAnnotation(CellStyle.class);
                } catch (NoSuchFieldException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }
}
