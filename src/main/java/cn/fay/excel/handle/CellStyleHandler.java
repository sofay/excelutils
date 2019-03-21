package cn.fay.excel.handle;

import cn.fay.excel.annotations.CellStyle;
import cn.fay.excel.annotations.ExcelExportField;
import cn.fay.excel.annotations.ExcelExportInfo;

/**
 * W: maybe org.apache.poi.ss.usermodel.Workbook
 * C: maybe org.apache.poi.ss.usermodel.CellStyle
 *
 * @author fay  fay9395@gmail.com
 * @date 2019-03-20 15:21.
 */
public interface CellStyleHandler<W, C> {
    /**
     * @param w               can generate c
     * @param commonCellStyle {@link ExcelExportInfo#defaultColumnNameCellStyle()} or {@link ExcelExportInfo#defaultColumnValueCellStyle()} ()}
     * @param fieldCellStyle  {@link ExcelExportField#columnNameCellStyle()} or {@link ExcelExportField#columnValueCellStyle()}
     * @param row
     * @param column
     * @param appendArgs
     * @return c
     */
    C handle(W w, CellStyle commonCellStyle, CellStyle fieldCellStyle, int row, int column, Object... appendArgs);
}
