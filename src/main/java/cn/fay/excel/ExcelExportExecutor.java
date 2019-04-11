package cn.fay.excel;

import cn.fay.excel.annotations.CellStyle;
import cn.fay.excel.annotations.ExcelExportField;
import cn.fay.excel.annotations.ExcelExportInfo;
import cn.fay.excel.annotations.ExcelExportModifyFields;
import cn.fay.excel.factory.ExcelObjectFactory;
import cn.fay.excel.handle.CellStyleHandler;
import cn.fay.excel.handle.DefaultExcelFieldTrans;
import cn.fay.excel.handle.ExcelFieldTrans;
import cn.fay.excel.util.PropertyLoader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 16:38.
 */
@SuppressWarnings("all")
public class ExcelExportExecutor {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelExportExecutor.class);
    private static ScriptEngine engine;
    private static Map<String, String> magicMethodMap;
    private static final String DEFAULT_MAGIC_METHOD_PROPERTIES = "magic.properties";
    private static final Pattern EVAL_PATTERN = Pattern.compile("`([ 0-9\\+\\-\\*/()\\$\\{\\}columnrow]*)`");
    private static final String[] INDEX_MAP = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
    private static final Integer CACHE_SIZE = 48;
    private static final Map<Class, Map<String, List<ExportFieldDescription>>> CLASS_EXCEL_HEAD_CACHE = new LinkedHashMap<Class, Map<String, List<ExportFieldDescription>>>(64, 0.75f, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry<Class, Map<String, List<ExportFieldDescription>>> eldest) {
            return size() > CACHE_SIZE;
        }
    };

    static {
        initMagicMethod(DEFAULT_MAGIC_METHOD_PROPERTIES);
    }

    private static void initMagicMethod(String propFile) {
        // load self magic method
        if (magicMethodMap == null) {
            magicMethodMap = new HashMap<>();
        }
        PropertyLoader.loader(propFile, new PropertyLoader.PropTravel() {
            @Override
            public void travel(String method, String eval) {
                if (magicMethodMap.containsKey(method)) {
                    throw new RuntimeException("find same magic method name: " + method);
                }
                magicMethodMap.put(method + "()", eval);
                if (LOGGER.isDebugEnabled()) {
                    LOGGER.debug("ExcelExportExecutor: load magic method: {} => {}", method, eval);
                }
            }
        });
    }

    public static void refreshMagicMethod(String propFile) {
        if (magicMethodMap != null) {
            magicMethodMap.clear();
        }
        initMagicMethod(propFile);
    }

    public static <E> Workbook excelWriter(List<E> data) {
        return excelWriter(data, true);
    }

    public static <E> Workbook excelWriter(List<E> data, boolean useLastRowValue) {
        return excelWriter(null, data, useLastRowValue, 0);
    }

    public static <E> Workbook excelWriter(Workbook workbook, List<E> data, boolean useLastRowValue, int sheetIndex) {
       return excelWriter(workbook, data, useLastRowValue, sheetIndex, null);
    }

    public static <E> Workbook excelWriter(Workbook workbook, List<E> data, boolean useLastRowValue, int sheetIndex, Class<E> clz) {
        Class clazz = clz;
        if(clazz == null) {
            if (data == null || data.isEmpty()) {
                return null;
            }
            clazz = data.get(0).getClass();
        }
        ExcelExportInfo excelExportInfo = (ExcelExportInfo) clazz.getAnnotation(ExcelExportInfo.class);
        assert excelExportInfo != null : String.format("%s 缺少 %s 注解", clazz.getName(), ExcelExportInfo.class.getSimpleName());

        String[] sheetNames = excelExportInfo.sheetNames();

        List<ExportFieldDescription> exportFieldDescriptions = collect(excelExportInfo.serialNo(), clazz);

        assert exportFieldDescriptions != null && !exportFieldDescriptions.isEmpty() : "导出时表头数据为空";
        workbook = workbook == null ? new SXSSFWorkbook(100) : workbook;
        assert sheetIndex >= 0 && sheetIndex <= sheetNames.length : String.format("sheet index must >= 0 and <= %s", sheetNames.length);
        createSheet(workbook, sheetNames[sheetIndex], collect(sheetIndex, exportFieldDescriptions), data, excelExportInfo, useLastRowValue);
        return workbook;
    }

    private static <E> void createSheet(Workbook wb, String sheetName, List<ExportFieldDescription> exportFieldDescriptions, List<E> data, ExcelExportInfo excelExportInfo, boolean useLastRowValue) {
        assert exportFieldDescriptions != null && !exportFieldDescriptions.isEmpty() : "请完成相关字段的注解填写";
        Sheet sheet = wb.getSheet(sheetName);
        sheet = sheet == null ? wb.createSheet(sheetName) : sheet;
        CellStyle defaultColumnNameCellStyle = excelExportInfo.defaultColumnNameCellStyle(); // column name style
        CellStyle defaultColumnValueCellStyle = excelExportInfo.defaultColumnValueCellStyle(); // column value style
        CellStyle defaultLastRowCellStyle = excelExportInfo.defaultLastRowCellStyle();
        int startRowIndex = sheet.getLastRowNum() + 1;
        if (startRowIndex == 1) {
            Row row = sheet.createRow(0);
            // column name
            for (int i = 0; i < exportFieldDescriptions.size(); i++) {
                ExportFieldDescription exportFieldDescription = exportFieldDescriptions.get(i);
                CellStyle fieldCS = exportFieldDescription.excelExportField.columnNameCellStyle(); // column name style
                Class<? extends CellStyleHandler> willUseHandleClass = chooseHandleClass(fieldCS, defaultColumnNameCellStyle);
                String columnName = exportFieldDescription.excelExportField.columnName();
                Cell cell = row.createCell(i);
                cell.setCellValue(columnName);
                cell.setCellType(fieldCS.cellType() != CellStyle.Value.Impl.defaultCellStyle().cellType() ? fieldCS.cellType() : defaultColumnNameCellStyle.cellType());
                cell.setCellStyle((org.apache.poi.ss.usermodel.CellStyle) ExcelObjectFactory.build(willUseHandleClass).handle(wb, defaultColumnNameCellStyle, fieldCS, 0, i));
                sheet.setColumnWidth(i, (columnName.getBytes().length + 5) * 256);
            }
        }

        if ((data == null || data.isEmpty()) && !useLastRowValue) {
            return;
        }
        // column value
        if (data != null) {
            for (int j = 0; j < data.size(); j++) {
                Row row = sheet.createRow(startRowIndex++);
                for (int i = 0; i < exportFieldDescriptions.size(); i++) {
                    Object instance = data.get(j);
                    Cell cell = row.createCell(i);
                    ExportFieldDescription exportFieldDescription = exportFieldDescriptions.get(i);
                    CellStyle tempFieldCS = exportFieldDescription.excelExportField.columnValueCellStyle();
                    CellStyle tempCommonCS = defaultColumnValueCellStyle;
                    Class<? extends CellStyleHandler> willUseHandleClass = chooseHandleClass(tempFieldCS, tempCommonCS);
                    //字段默认值
                    String defaultValue = exportFieldDescription.excelExportField.defaultValue();
                    //字段值
                    Object value = null;
                    boolean useDefault = false;
                    try {
                        value = exportFieldDescription.field.get(instance);
                    } catch (Exception e) {
                        // ignore
                        e.printStackTrace();
                    }
                    if (value == null) {
                        value = handleDefaultValuePlaceHolder(handleSelfMagicMethod(defaultValue), cell.getRow().getRowNum(), cell.getColumnIndex());
                        useDefault = true;
                    }
                    if (exportFieldDescription.transClass != null && !useDefault) {
                        value = ExcelObjectFactory.build(exportFieldDescription.transClass).trans(value);
                    }
                    cell.setCellType(tempFieldCS.cellType() != CellStyle.Value.Impl.defaultCellStyle().cellType() ? tempFieldCS.cellType() : tempCommonCS.cellType());
                    cell.setCellStyle((org.apache.poi.ss.usermodel.CellStyle) ExcelObjectFactory.build(willUseHandleClass).handle(wb, tempCommonCS, tempFieldCS, row.getRowNum(), i));
                    if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        String val = String.valueOf(value);
                        cell.setCellValue("".equals(val) ? 0 : Double.valueOf(val));
                    } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        cell.setCellFormula(String.valueOf(value));
                    } else {
                        cell.setCellValue(String.valueOf(value));
                    }
                }
            }
        }
        if (useLastRowValue)
            handleLastRow(wb, sheet.createRow(startRowIndex++), exportFieldDescriptions, defaultLastRowCellStyle);
    }

    private static void handleLastRow(Workbook wb, Row row, List<ExportFieldDescription> exportFieldDescriptions, CellStyle defaultLastRowCellStyle) {
        for (int i = 0; i < exportFieldDescriptions.size(); i++) {
            Cell cell = row.createCell(i);
            ExportFieldDescription exportFieldDescription = exportFieldDescriptions.get(i);
            CellStyle tempFieldCS = exportFieldDescription.excelExportField.lastRowCellStyle();
            CellStyle tempCommonCS = defaultLastRowCellStyle;
            Class<? extends CellStyleHandler> willUseHandleClass = chooseHandleClass(tempFieldCS, tempCommonCS);
            //字段默认值
            String defaultValue = exportFieldDescription.excelExportField.lastRowValue();
            //字段值
            Object value = null;
            value = handleDefaultValuePlaceHolder(handleSelfMagicMethod(defaultValue), cell.getRow().getRowNum(), cell.getColumnIndex());
            cell.setCellType(tempFieldCS.cellType() != CellStyle.Value.Impl.defaultCellStyle().cellType() ? tempFieldCS.cellType() : tempCommonCS.cellType());
            cell.setCellStyle((org.apache.poi.ss.usermodel.CellStyle) ExcelObjectFactory.build(willUseHandleClass).handle(wb, tempCommonCS, tempFieldCS, row.getRowNum(), i));
            if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                String val = String.valueOf(value);
                cell.setCellValue("".equals(val) ? 0 : Double.valueOf(val));
            } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                cell.setCellFormula(String.valueOf(value));
            } else {
                cell.setCellValue(String.valueOf(value));
            }
        }
    }

    private static String handleSelfMagicMethod(String magicMethodName) {
        if (magicMethodMap != null && magicMethodMap.containsKey(magicMethodName)) {
            return magicMethodMap.get(magicMethodName);
        }
        return magicMethodName;
    }

    private static String handleDefaultValuePlaceHolder(String defaultValue, Object... args) {
        assert defaultValue != null && args.length > 2;
        Matcher matcher = EVAL_PATTERN.matcher(defaultValue);
        while (matcher.find()) {
            if (engine == null) {
                engine = new ScriptEngineManager().getEngineByName("js");
            }
            String expr = matcher.group(1);
            boolean isColumnExpr = expr.contains("${column}");
            boolean isRowExpr = expr.contains("${row}");
            if (isColumnExpr) {
                expr = expr.replaceAll("\\$\\{column}", String.valueOf(args[1]));
            }
            if (isRowExpr) {
                expr = expr.replaceAll("\\$\\{row}", toRowName((int) args[0]));
            }
            try {
                Object ret = engine.eval(expr);
                defaultValue = defaultValue.replace(matcher.group(), isColumnExpr ? toColumnName(Integer.parseInt(String.valueOf(ret))) : String.valueOf(ret));
            } catch (ScriptException e) {
                throw new RuntimeException(e);
            }
        }
        return "".equals(defaultValue) ? defaultValue : defaultValue.replaceAll("\\$\\{row}", toRowName((int) args[0])).replaceAll("\\$\\{column}", toColumnName((int) args[1]));
    }

    private static String toColumnName(int index) {
        int h = index / 26;
        int l = index % 26;
        return h <= 0 ? INDEX_MAP[l] : INDEX_MAP[h - 1] + INDEX_MAP[l];
    }

    private static String toRowName(int index) {
        return String.valueOf(index + 1);
    }

    private static List<ExportFieldDescription> collect(int sheetIndex, List<ExportFieldDescription> descriptions) {
        assert descriptions != null;
        List<ExportFieldDescription> result = new ArrayList<>(descriptions.size() < 10 ? descriptions.size() : 10);
        for (ExportFieldDescription description : descriptions) {
            if (description.usedForSheet.equals(description.usedForAllSheets) || description.usedForSheet.contains(sheetIndex)) {
                result.add(description);
            }
        }
        return result;
    }

    private static List<ExportFieldDescription> collect(String serialNo, Class cls) {
        assert serialNo != null;
        List<String> excludeFields = getExcludeFields(cls);
        Map<String, List<ExportFieldDescription>> map = CLASS_EXCEL_HEAD_CACHE.get(cls);
        List<ExportFieldDescription> result = map == null ? null : map.get(serialNo);
        if (result == null) {
            synchronized (CLASS_EXCEL_HEAD_CACHE) {
                map = CLASS_EXCEL_HEAD_CACHE.get(cls);
                if (map == null) {
                    map = new HashMap<>();
                }
                if ((result = map.get(serialNo)) != null) {
                    return result;
                }
                List<Field> fields = getFields(cls, Object.class);
                result = new ArrayList<>();
                for (Field field : fields) {
                    if (excludeFields != null && excludeFields.contains(field.getName())) {
                        continue;
                    }
                    ExcelExportField excelExportField = field.getAnnotation(ExcelExportField.class);
                    if (excelExportField != null && serialNo.equals(excelExportField.usedForSerialNo())) { // suit serial no
                        result.add(new ExportFieldDescription(field, excelExportField));
                    }
                }
                assert result != null && !result.isEmpty() : String.format("导出类 %s 没有字段有 %s 注解", cls.getSimpleName(), ExcelExportField.class.getSimpleName());
                Collections.sort(result);
                map.put(serialNo, result);
                CLASS_EXCEL_HEAD_CACHE.put(cls, map);
            }
        }
        return result;
    }

    private static Class<? extends CellStyleHandler> chooseHandleClass(CellStyle fieldCS, CellStyle commonCS) {
        // if field handle class not equals default handle class then return field handle class
        // else return common handle class (not need compare with default handle class)
        Class<? extends CellStyleHandler> fieldClass = chooseHandleClassFromSingleAnnotation(fieldCS);
        return !fieldClass.equals(CellStyle.Value.Impl.defaultCellStyle().cellStyleHandlerClass()) ? fieldClass : chooseHandleClassFromSingleAnnotation(commonCS);
    }

    private static Class<? extends CellStyleHandler> chooseHandleClassFromSingleAnnotation(CellStyle cellStyle) {
        try {
            return !cellStyle.cellStyleHandlerClass().equals(CellStyle.Value.Impl.defaultCellStyle().cellStyleHandlerClass()) ?
                    cellStyle.cellStyleHandlerClass() :
                    (Class<? extends CellStyleHandler>) (!cellStyle.cellStyleHandlerClassName().equals(CellStyle.Value.Impl.defaultCellStyle().cellStyleHandlerClassName()) ?
                            Class.forName(cellStyle.cellStyleHandlerClassName()) : cellStyle.cellStyleHandlerClass());
        } catch (ClassNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    /*private static String getExportFileName(Class targetCls) {
        ExcelExportInfo excelExportInfo = (ExcelExportInfo) targetCls.getAnnotation(ExcelExportInfo.class);
        assert excelExportInfo != null : String.format("%s 缺少 %s 注解", targetCls.getName(), ExcelExportInfo.class.getSimpleName());
        String prefix = excelExportInfo.fileNamePrefix();
        String className = excelExportInfo.fileNameSuffixGenerateClassName();
        Class cls = excelExportInfo.fileNameSuffixGenerateClass();
        Class defaultGenerateClass = DefaultExcelExportFileNameSuffixGenerate.class;
        Class<? extends ExcelExportFileNameSuffixGenerate> suffixGenerateClass = defaultGenerateClass;
        // 有一个不是默认值
        if (!defaultGenerateClass.getName().equals(className) || !defaultGenerateClass.equals(cls)) {
            try {
                Class unChecked = defaultGenerateClass.equals(cls) ? Class.forName(className) : cls;
                if (ExcelExportFileNameSuffixGenerate.class.isAssignableFrom(unChecked)) {
                    suffixGenerateClass = unChecked;
                } else {
                    throw new IllegalArgumentException(String.format("%s 应该要实现 %s", unChecked.getName(), ExcelFieldTrans.class.getName()));
                }
            } catch (ClassNotFoundException e) {
                throw new IllegalArgumentException(String.format("请检查 %s 中  ExcelExportInfo 注解中 fileNameSuffixGenerateClassName 设置的值", targetCls.getSimpleName()));
            }
        }
        String suffix = ExcelExportFileNameSuffixGenerateFactory.build(suffixGenerateClass).generateSuffix(prefix);
        return prefix + suffix;
    }*/

    private static List<String> getExcludeFields(Class targetCls) {
        ExcelExportModifyFields excelExportModifyFields = (ExcelExportModifyFields) targetCls.getAnnotation(ExcelExportModifyFields.class);
        if (excelExportModifyFields != null) {
            return Arrays.asList(excelExportModifyFields.excludeFields());
        }
        return Collections.emptyList();
    }

    private static List<Field> getFields(Class<?> clazz, Class<?> stopClass) {
        try {
            List<Field> fieldList = new ArrayList<>();
            while (clazz != null && clazz != stopClass) {//当父类为null的时候说明到达了最上层的父类(Object类).
                fieldList.addAll(Arrays.asList(clazz.getDeclaredFields()));
                clazz = clazz.getSuperclass(); //得到父类,然后赋给自己
            }
            return fieldList;
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
    }


    private static class ExportFieldDescription implements Comparable<ExportFieldDescription> {
        Class<? extends ExcelFieldTrans> transClass;
        Field field;
        ExcelExportField excelExportField;
        static Set<Integer> usedForAllSheets = new HashSet<Integer>() {{add(-1);}};
        Set<Integer> usedForSheet;

        ExportFieldDescription(Field field, ExcelExportField excelExportField) {
            assert field != null;
            assert excelExportField != null;
            String transClassName = excelExportField.transClassName();
            Class cls = excelExportField.transClass();
            Class defaultTransClass = DefaultExcelFieldTrans.class;
            // 有一个不是默认值
            if (!defaultTransClass.getName().equals(transClassName) || !defaultTransClass.equals(cls)) {
                try {
                    Class unChecked = defaultTransClass.equals(cls) ? Class.forName(transClassName) : cls;
                    if (ExcelFieldTrans.class.isAssignableFrom(unChecked)) {
                        transClass = unChecked;
                    } else {
                        throw new IllegalArgumentException(String.format("%s 应该要实现 %s", unChecked.getName(), ExcelFieldTrans.class.getName()));
                    }
                } catch (ClassNotFoundException e) {
                    throw new IllegalArgumentException(String.format("请检查 %s 中 %s 字段的 ExcelExportField 注解中 transClassName 设置的值", field.getDeclaringClass().getSimpleName(), field.getName()));
                }
            }
            this.field = field;
            if (!this.field.isAccessible()) {
                this.field.setAccessible(true);
            }
            this.excelExportField = excelExportField;
            usedForSheet = new HashSet<>();
            for (int index : excelExportField.usedForSheetIndex()) {
                usedForSheet.add(index);
            }
        }

        @Override
        public int compareTo(ExportFieldDescription o) {
            double order1 = excelExportField == null ? 9999 : excelExportField.order(); // null 的排在前面排在后面都一样
            double order2 = o == null ? 9999 : o.excelExportField.order();
            return (int) (order1 - order2);
        }
    }
}
