package cn.fay.excel.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.math.BigDecimal;

/**
 * 用于fast json 序列化时保留小数用
 * {@see com.raycloud.kmsy.common.fastjson.DecimalDescValueFilter}
 *
 * @author fay  fay9395@gmail.com
 * @date 2018/11/15 10:32 AM.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface DecimalDesc {
    /**
     * 填写需要保留的小数位
     */
    int scale() default 3;

    /**
     * 不够位数时是否在后面补0
     */
    boolean zeroAppend() default false;

    /**
     * {@link BigDecimal#ROUND_HALF_UP}
     */
    int roundingMode() default BigDecimal.ROUND_HALF_UP;
}
