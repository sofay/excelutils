package cn.fay.excel.factory;


import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author fay  fay9395@gmail.com
 * @date 2018/8/9 下午5:02.
 */
public class ExcelObjectFactory {
    private static final Map<Class, Object> CACHE = new ConcurrentHashMap<>();


    public static <T> T build(Class<T> cls) {
        Object obj = CACHE.get(cls);
        if (obj == null) {
            synchronized (CACHE) {
                try {
                    obj = cls.getConstructor().newInstance();
                    CACHE.put(cls, obj);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return (T) obj;
    }
}
