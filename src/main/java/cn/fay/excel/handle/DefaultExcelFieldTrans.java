package cn.fay.excel.handle;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午5:00.
 */
public class DefaultExcelFieldTrans extends AbstractExcelFieldTrans {
    @Override
    public Object doTrans(Object argv) {
        return argv;
    }
}
