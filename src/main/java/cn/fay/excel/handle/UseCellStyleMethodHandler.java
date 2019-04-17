package cn.fay.excel.handle;

import cn.fay.excel.config.DefaultValueConfigContext;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-20 15:20.
 */
public class UseCellStyleMethodHandler implements CellStyleHandler<Workbook, CellStyle> {
    private static final Logger LOGGER = LoggerFactory.getLogger(UseCellStyleMethodHandler.class);
    private static final Method[] ANNOTATION_CELL_STYLE_METHODS;
    private static final Method[] ANNOTATION_FONT_METHODS;
    private static final Set<String> ANNOTATION_EXCLUDE_METHOD_NAMES = new HashSet<>();
    private static final Map<String, Method> POI_CELL_STYLE_METHODS = new HashMap<>();
    private static final Map<String, Method> POI_FONT_METHODS = new HashMap<>();
    private static final Integer CACHE_SIZE = 64;
    private static final LRUCache<String, CellStyle> CELL_STYLE_LRU_CACHE = new LRUCache<>(CACHE_SIZE);
    private static final LRUCache<String, Font> FONT_LRU_CACHE = new LRUCache<>(CACHE_SIZE);


    static {
        ANNOTATION_EXCLUDE_METHOD_NAMES.add("cellStyleHandlerClassName");
        ANNOTATION_EXCLUDE_METHOD_NAMES.add("cellStyleHandlerClass");
        ANNOTATION_EXCLUDE_METHOD_NAMES.add("cellType");
        Method[] methods = cn.fay.excel.annotations.CellStyle.class.getDeclaredMethods();
        ANNOTATION_CELL_STYLE_METHODS = new Method[methods.length - ANNOTATION_EXCLUDE_METHOD_NAMES.size()];
        int index = 0;
        for (Method method : methods) {
            if (ANNOTATION_EXCLUDE_METHOD_NAMES.contains(method.getName())) {
                continue;
            }
            ANNOTATION_CELL_STYLE_METHODS[index++] = method;
            if (!method.isAccessible()) {
                method.setAccessible(true);
            }
        }
        for (Method method : CellStyle.class.getDeclaredMethods()) {
            if (method.getName().startsWith("set")) {
                if (!method.isAccessible()) {
                    method.setAccessible(true);
                }
                POI_CELL_STYLE_METHODS.put(method.getName(), method);
            }
        }
        ANNOTATION_FONT_METHODS = cn.fay.excel.annotations.Font.class.getDeclaredMethods();
        for (Method method : ANNOTATION_FONT_METHODS) {
            if (!method.isAccessible()) {
                method.setAccessible(true);
            }
        }
        for (Method method : Font.class.getDeclaredMethods()) {
            if (method.getName().startsWith("set")) {
                if (!method.isAccessible()) {
                    method.setAccessible(true);
                }
                POI_FONT_METHODS.put(method.getName(), method);
            }
        }
    }

    @Override
    public CellStyle handle(Workbook workbook, cn.fay.excel.annotations.CellStyle commonCellStyle, cn.fay.excel.annotations.CellStyle fieldCellStyle, int row, int column, Object... appendArgs) {
        CellStyle ret;
        String cacheKey = getCacheKey(workbook, commonCellStyle, fieldCellStyle);
        if ((ret = CELL_STYLE_LRU_CACHE.get(cacheKey)) == null) {
            CELL_STYLE_LRU_CACHE.put(cacheKey, ret = workbook.createCellStyle());
            init(workbook, commonCellStyle, fieldCellStyle, ret);
        }
        return ret;
    }

    private String getCacheKey(Workbook workbook, cn.fay.excel.annotations.CellStyle commonCellStyle, cn.fay.excel.annotations.CellStyle fieldCellStyle) {
        StringBuilder key = new StringBuilder();
        for (Method method : ANNOTATION_CELL_STYLE_METHODS) {
            try {
                key.append(method.invoke(workbook)).append(method.invoke(commonCellStyle))
                        .append(method.invoke(fieldCellStyle));
            } catch (IllegalAccessException | InvocationTargetException e) {
                LOGGER.error("UseCellStyleMethodHandler cell style:", e);
                return null;
            }
        }
        return key.toString();
    }

    private String getCacheKey(Workbook workbook, cn.fay.excel.annotations.Font commonFont, cn.fay.excel.annotations.Font fieldFont) {
        StringBuilder key = new StringBuilder();
        for (Method method : ANNOTATION_FONT_METHODS) {
            try {
                key.append(method.invoke(workbook)).append(method.invoke(commonFont))
                        .append(method.invoke(fieldFont));
            } catch (IllegalAccessException | InvocationTargetException e) {
                LOGGER.error("UseCellStyleMethodHandler font:", e);
                return null;
            }
        }
        return key.toString();
    }

    private void init(Workbook workBook, cn.fay.excel.annotations.CellStyle commonCellStyle, cn.fay.excel.annotations.CellStyle fieldCellStyle, CellStyle cellStyle) {
        for (Method method : ANNOTATION_CELL_STYLE_METHODS) {
            try {
                Object val = method.invoke(fieldCellStyle);
                Object defaultVal = method.invoke(cn.fay.excel.annotations.CellStyle.Value.Impl.defaultCellStyle());
                Object commonVal = method.invoke(commonCellStyle);
                Object transfedVal = null;
                if ((transfedVal = calcValue(method, defaultVal, val, commonVal)) != null) {
                    String mayMethodName = "set" + method.getName().substring(0, 1).toUpperCase() + method.getName().substring(1);
                    if (!POI_CELL_STYLE_METHODS.containsKey(mayMethodName)) {
                        LOGGER.warn("UseCellStyleMethodHandler can not resolve method: {}", mayMethodName);
                        continue;
                    }
                    switch (mayMethodName) {
                        case "setFont":
                            cellStyle.setFont(fontMethod(workBook, commonCellStyle, fieldCellStyle));
                            break;
                        default:
                            POI_CELL_STYLE_METHODS.get(mayMethodName).invoke(cellStyle, transfedVal);
                    }
                }
            } catch (IllegalAccessException | InvocationTargetException e) {
                LOGGER.error("UseCellStyleMethodHandler invoke method: " + method.getName() + " error.", e);
            }
        }
    }

    private Object calcValue(Method method, Object defaultVal, Object fieldVal, Object commonVal) {
        if (!fieldVal.equals(defaultVal)) {
            return fieldVal;
        }
        if (!commonVal.equals(defaultVal)) {
            return commonVal;
        }
        String transfedVal = DefaultValueConfigContext.transDefaultValue(getPropKey(method));
        if (transfedVal != null) {
            if (method.getReturnType().equals(short.class)) {
                return Short.parseShort(transfedVal);
            }
            if (method.getReturnType().equals(int.class)) {
                return Integer.parseInt(transfedVal);
            }
            return transfedVal;
        }else {
            return defaultVal;
        }
    }

    private String getPropKey(Method method) {
        return method.getDeclaringClass().getSimpleName().toLowerCase() + "." + method.getName().toLowerCase();
    }

    private Font fontMethod(Workbook workBook, cn.fay.excel.annotations.CellStyle commonCellStyle, cn.fay.excel.annotations.CellStyle fieldCellStyle) throws InvocationTargetException, IllegalAccessException {
        cn.fay.excel.annotations.Font defFont = cn.fay.excel.annotations.CellStyle.Value.Impl.defaultCellStyle().font();
        cn.fay.excel.annotations.Font commFont = commonCellStyle.font();
        cn.fay.excel.annotations.Font fieldFont = fieldCellStyle.font();
        Font font;
        String cacheKey = getCacheKey(workBook, commFont, fieldFont);
        if ((font = FONT_LRU_CACHE.get(cacheKey)) == null) {
            font = workBook.createFont();
            for (Method fontMethod : ANNOTATION_FONT_METHODS) {
                Object fontFieldVal = fontMethod.invoke(fieldFont);
                Object fontDefVal = fontMethod.invoke(defFont);
                Object fontCommVal = fontMethod.invoke(commFont);
                Object transfedVal = null;
                if ((transfedVal = calcValue(fontMethod, fontDefVal, fontFieldVal, fontCommVal)) != null) {
                    String mayFontMethodName = "set" + fontMethod.getName().substring(0, 1).toUpperCase() + fontMethod.getName().substring(1);
                    if (!POI_FONT_METHODS.containsKey(mayFontMethodName)) {
                        LOGGER.warn("UseCellStyleMethodHandler can not resolve font method: {}", mayFontMethodName);
                        continue;
                    }
                    POI_FONT_METHODS.get(mayFontMethodName).invoke(font, transfedVal);
                }
            }
            FONT_LRU_CACHE.put(cacheKey, font);
        }
        return font;
    }


    static class LRUCache<K, V> {
        private LinkedHashMap<K, V> cache;

        public LRUCache(final int capacity) {
            cache = new LinkedHashMap<K, V>(capacity, 0.75f, true) {
                @Override
                protected boolean removeEldestEntry(Map.Entry eldest) {
                    if (size() > capacity) {
                        LOGGER.info("LRC remove cache:" + eldest.getKey());
                        return true;
                    }
                    return super.removeEldestEntry(eldest);
                }
            };
        }

        public V get(K k) {
            return cache.get(k);
        }

        public void put(K k, V v) {
            cache.put(k, v);
        }
    }
}
