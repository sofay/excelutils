package cn.fay.excel.annotations;

import cn.fay.excel.handle.CellStyleHandler;
import cn.fay.excel.handle.UseCellStyleMethodHandler;

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
    String STRING_DEFAULT_VALUE = "DEFAULT";
    short SHORT_DEFAULT_VALUE = 0;

    Font font() default @Font;

    short fontHeightInPoints() default SHORT_DEFAULT_VALUE;

    short alignment() default SHORT_DEFAULT_VALUE;

    short borderBottom() default SHORT_DEFAULT_VALUE;

    short borderLeft() default SHORT_DEFAULT_VALUE;

    short borderRight() default SHORT_DEFAULT_VALUE;

    short borderTop() default SHORT_DEFAULT_VALUE;

    short fillPattern() default SHORT_DEFAULT_VALUE;

    short fillForegroundColor() default SHORT_DEFAULT_VALUE;

    short dataFormat() default SHORT_DEFAULT_VALUE;

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
                try {
                    return defaultCS = Value.class.getField("DEFAULT_CELL_STYLE").getAnnotation(CellStyle.class);
                } catch (NoSuchFieldException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }
}
