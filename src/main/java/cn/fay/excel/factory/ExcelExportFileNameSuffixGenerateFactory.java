package cn.fay.excel.factory;

import cn.fay.excel.handle.ExcelExportFileNameSuffixGenerate;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午5:02.
 */
public class ExcelExportFileNameSuffixGenerateFactory {
    private static final Map<Class<? extends ExcelExportFileNameSuffixGenerate>, ExcelExportFileNameSuffixGenerate> CACHE = new ConcurrentHashMap<>();

    /**
     * 不用刻意去维持线程安全
     */
    public static ExcelExportFileNameSuffixGenerate build(Class<? extends ExcelExportFileNameSuffixGenerate> cls) {
        ExcelExportFileNameSuffixGenerate obj = CACHE.get(cls);
        if (obj == null) {
            try {
                obj = cls.getConstructor().newInstance();
                CACHE.put(cls, cls.getConstructor().newInstance());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return obj;
    }
}
