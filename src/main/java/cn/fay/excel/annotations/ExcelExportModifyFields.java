package cn.fay.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导出时支持修改某些字段，比如在某个父类定义了一个 {@link ExcelExportField} 但是在某个子类导出时不需要这列数据，就可以在子类使用该注解排除父类的字段
 *
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午6:53.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Deprecated
public @interface ExcelExportModifyFields {
    /**
     * 要排除的列对应的字段名 e.g. {@code @ExcelExportModifyFields(excludeFields = "created")}
     */
    String[] excludeFields() default {};

    /**
     * 子类覆盖父类中 {@link ExcelExportField} 中的设置
     * 目前不打算实现，可以自行实现
     */
}
