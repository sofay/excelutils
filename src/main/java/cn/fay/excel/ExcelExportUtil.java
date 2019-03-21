package cn.fay.excel;

import cn.fay.excel.annotations.CellStyle;
import cn.fay.excel.annotations.ExcelExportField;
import cn.fay.excel.annotations.ExcelExportInfo;
import cn.fay.excel.annotations.ExcelExportModifyFields;
import cn.fay.excel.factory.ExcelExportFileNameSuffixGenerateFactory;
import cn.fay.excel.factory.ExcelObjectFactory;
import cn.fay.excel.handle.CellStyleHandler;
import cn.fay.excel.handle.DefaultExcelExportFileNameSuffixGenerate;
import cn.fay.excel.handle.DefaultExcelFieldTrans;
import cn.fay.excel.handle.ExcelExportFileNameSuffixGenerate;
import cn.fay.excel.handle.ExcelFieldTrans;
import com.sun.istack.internal.Nullable;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletResponse;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 16:38.
 */
@SuppressWarnings("all")
public class ExcelExportUtil {
    private static final Integer CACHE_SIZE = 48;
    private static final Map<Class, List<ExportFieldDescription>> CLASS_EXCEL_HEAD_CACHE = new LinkedHashMap<Class, List<ExportFieldDescription>>(64, 0.75f, true) {
        @Override
        protected boolean removeEldestEntry(Map.Entry<Class, List<ExportFieldDescription>> eldest) {
            return size() > CACHE_SIZE;
        }
    };


    public static <E> void excelWriter(List<E> data, ServletResponse response) throws Exception {
        Class clazz = data.get(0).getClass();
        ExcelExportInfo excelExportInfo = (ExcelExportInfo) clazz.getAnnotation(ExcelExportInfo.class);
        assert excelExportInfo != null : String.format("%s 缺少 %s 注解", clazz.getName(), ExcelExportInfo.class.getSimpleName());
        Workbook workbook = excelWriter(data);
        String fileName = getExportFileName(clazz);
        //导出数据
        export(fileName, workbook, (HttpServletResponse) response);
    }

    public static <E> Workbook excelWriter(List<E> data) {
        if (data == null || data.isEmpty()) {
            return null;
        }
        Class clazz = data.get(0).getClass();
        ExcelExportInfo excelExportInfo = (ExcelExportInfo) clazz.getAnnotation(ExcelExportInfo.class);
        assert excelExportInfo != null : String.format("%s 缺少 %s 注解", clazz.getName(), ExcelExportInfo.class.getSimpleName());

        String sheetName = excelExportInfo.sheetName();

        List<ExportFieldDescription> exportFieldDescriptions = getExportFieldDescriptionList(clazz);

        assert exportFieldDescriptions != null && !exportFieldDescriptions.isEmpty() : "导出时表头数据为空";
        SXSSFWorkbook wb = new SXSSFWorkbook(100);
        createSheet(wb, sheetName, exportFieldDescriptions, data, excelExportInfo);
        return wb;
    }

    private static <E> void createSheet(Workbook wb, String sheetName, List<ExportFieldDescription> exportFieldDescriptions, List<E> data, ExcelExportInfo excelExportInfo) {
        assert exportFieldDescriptions != null && !exportFieldDescriptions.isEmpty() : "请完成相关字段的注解填写";
        Sheet sheet = wb.createSheet(sheetName);
        Row row = sheet.createRow(0);
        CellStyle defaultColumnNameCellStyle = excelExportInfo.defaultColumnNameCellStyle(); // column name style
        CellStyle defaultColumnValueCellStyle = excelExportInfo.defaultColumnValueCellStyle(); // column value style
        CellStyle defaultLastRowCellStyle = excelExportInfo.defaultLastRowCellStyle();

        // column name
        for (int i = 0; i < exportFieldDescriptions.size(); i++) {
            ExportFieldDescription exportFieldDescription = exportFieldDescriptions.get(i);
            CellStyle fieldCS = exportFieldDescription.excelExportField.columnNameCellStyle(); // column name style
            Class<? extends CellStyleHandler> willUseHandleClass = chooseHandleClass(fieldCS, defaultColumnNameCellStyle);
            String columnName = exportFieldDescription.excelExportField.columnName();
            Cell cell = row.createCell(i);
            cell.setCellValue(columnName);
            cell.setCellType(exportFieldDescription.excelExportField.columnNameCellType());
            cell.setCellStyle((org.apache.poi.ss.usermodel.CellStyle) ExcelObjectFactory.build(willUseHandleClass).handle(wb, defaultColumnNameCellStyle, fieldCS, 0, i));
            sheet.setColumnWidth(i, (columnName.getBytes().length + 5) * 256);
        }
        if (data == null || data.isEmpty()) {
            return;
        }
        // column value
        for (int j = 0; j < data.size(); j++) {
            Object instance = data.get(j);
            row = sheet.createRow(j + 1);
            boolean isLastRow = j + 1 >= data.size();
            for (int i = 0; i < exportFieldDescriptions.size(); i++) {
                Cell cell = row.createCell(i);
                ExportFieldDescription exportFieldDescription = exportFieldDescriptions.get(i);
                CellStyle tempFieldCS = isLastRow ? exportFieldDescription.excelExportField.lastRowCellStyle() : exportFieldDescription.excelExportField.columnValueCellStyle();
                CellStyle tempCommonCS = isLastRow ? defaultLastRowCellStyle : defaultColumnValueCellStyle;
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
                    value = defaultValue;
                    useDefault = true;
                }
                if (exportFieldDescription.transClass != null && !useDefault) {
                    value = ExcelObjectFactory.build(exportFieldDescription.transClass).trans(value);
                }
                cell.setCellType(exportFieldDescription.excelExportField.columnValueCellType());
                cell.setCellStyle((org.apache.poi.ss.usermodel.CellStyle) ExcelObjectFactory.build(willUseHandleClass).handle(wb, tempCommonCS, tempFieldCS, row.getRowNum(), i));
                cell.setCellValue(value.toString());
            }
        }
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

    private static void export(String fileName, Workbook wb, HttpServletResponse response) throws Exception {
        // 清空response
        response.reset();

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddhh");

        fileName += "_" + simpleDateFormat.format(new Date());

        // 编码处理，对于linux 操作系统 （ linux默认 utf-8,windows 默认 GBK)
        String defaultEncoding = System.getProperty("file.encoding");
        //匹配文件本事与名称的版本
        String endName = ".xlsx";
        if (wb instanceof HSSFWorkbook) {
            endName = ".xls";
        }
        if (defaultEncoding != null && defaultEncoding.equals("UTF-8")) {
            response.addHeader("Content-Disposition", "attachment;filename="
                    + new String(fileName.getBytes("GBK"), "iso-8859-1") + endName);
        } else {
            response.addHeader("Content-Disposition", "attachment;filename="
                    + new String(fileName.getBytes(), "iso-8859-1") + endName);
        }
        response.setCharacterEncoding("utf-8");
        // 设置response的Header
        response.setContentType("application/vnd.ms-excel");
        OutputStream ouputStream = response.getOutputStream();
        wb.write(ouputStream);
        ouputStream.flush();
        ouputStream.close();
    }


    private static String getExportFileName(Class targetCls) {
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
    }

    private static List<String> getExcludeFields(Class targetCls) {
        ExcelExportModifyFields excelExportModifyFields = (ExcelExportModifyFields) targetCls.getAnnotation(ExcelExportModifyFields.class);
        if (excelExportModifyFields != null) {
            return Arrays.asList(excelExportModifyFields.excludeFields());
        }
        return Collections.emptyList();
    }

    private static List<ExportFieldDescription> getExportFieldDescriptionList(Class<?> cls) {
        List<String> excludeFields = getExcludeFields(cls);
        List<ExportFieldDescription> result = CLASS_EXCEL_HEAD_CACHE.get(cls);
        if (result == null) {
            synchronized (CLASS_EXCEL_HEAD_CACHE) {
                result = CLASS_EXCEL_HEAD_CACHE.get(cls);
                if (result != null) {
                    return result;
                }
                List<Field> fields = getFields(cls, Object.class);
                result = new ArrayList<>();
                for (Field field : fields) {
                    if (excludeFields != null && excludeFields.contains(field.getName())) {
                        continue;
                    }
                    ExcelExportField excelExportField = field.getAnnotation(ExcelExportField.class);
                    if (excelExportField != null) {
                        result.add(new ExportFieldDescription(field, excelExportField));
                    }
                }
                assert result != null && !result.isEmpty() : String.format("导出类 %s 没有字段有 %s 注解", cls.getSimpleName(), ExcelExportField.class.getSimpleName());
                Collections.sort(result);
                CLASS_EXCEL_HEAD_CACHE.put(cls, result);
            }
        }
        return result;
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
        @Nullable
        Class<? extends ExcelFieldTrans> transClass;
        Field field;
        ExcelExportField excelExportField;

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
        }

        @Override
        public int compareTo(ExportFieldDescription o) {
            double order1 = excelExportField == null ? 9999 : excelExportField.order(); // null 的排在前面排在后面都一样
            double order2 = o == null ? 9999 : o.excelExportField.order();
            return (int) (order1 - order2);
        }
    }
}
