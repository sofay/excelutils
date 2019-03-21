package cn.fay.excel.handle;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午4:57.
 */
public interface ExcelFieldTrans<R, T> {
    R trans(T argv);
}
