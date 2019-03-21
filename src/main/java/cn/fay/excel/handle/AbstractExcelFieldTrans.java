package cn.fay.excel.handle;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午5:39.
 */
public abstract class AbstractExcelFieldTrans<R, T> implements ExcelFieldTrans<R,T> {

    public R trans(T argv) {
        if (argv != null) {
            return doTrans(argv);
        }
        return null;
    }

    public abstract R doTrans(T argv);
}
