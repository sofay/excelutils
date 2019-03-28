package cn.fay.test;

import cn.fay.excel.ExcelExportExecutor;
import org.junit.Assert;
import org.junit.Test;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-28 19:41.
 */
public class PropLoaderTest {

    @Test
    public void magicMethodTest() throws Exception {
        Field field = ExcelExportExecutor.class.getDeclaredField("magicMethodMap");
        if (!field.isAccessible()) {
            field.setAccessible(true);
        }
        Map<String, String> map = new HashMap<>();
        map.put("faySum()", "sum(${column}2:${column}`${row} - 1`)");
        map.put("magic1()", "`${column} - 2`${row}/`${column} - 1`${row}");
        map.put("magic2()", "`${column} - 3`${row}/max(`${column} - 3`:`${column} - 3`)");
        Assert.assertTrue(map.equals(field.get(null)));
    }
}
