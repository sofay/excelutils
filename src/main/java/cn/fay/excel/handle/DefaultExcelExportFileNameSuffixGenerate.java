package cn.fay.excel.handle;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午8:13.
 */
public class DefaultExcelExportFileNameSuffixGenerate implements ExcelExportFileNameSuffixGenerate {
    private static final SimpleDateFormat SIMPLE_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HHmmss");

    @Override
    public String generateSuffix(String prefix) {
        return SIMPLE_DATE_FORMAT.format(new Date());
    }
}
