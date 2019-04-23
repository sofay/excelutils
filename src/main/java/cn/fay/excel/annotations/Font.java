package cn.fay.excel.annotations;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 11:06.
 */
public @interface Font {
    String fontName() default CellStyle.STRING_DEFAULT_VALUE;
//    fontHeight() default CellStyle.;
//    italic() default CellStyle.;
//    strikeout() default CellStyle.;
//    color() default CellStyle.;
//    typeOffset() default CellStyle.;
//    underline() default CellStyle.;
//    charSet() default CellStyle.;
//    charSet() default CellStyle.;
    short fontHeightInPoints() default CellStyle.SHORT_DEFAULT_VALUE_10;

    short boldweight() default CellStyle.SHORT_DEFAULT_VALUE;
}
