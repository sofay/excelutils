package cn.fay.test;

import cn.fay.excel.ExcelExportExecutor;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import java.lang.reflect.Field;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-21 19:57.
 */
public class EVALTest {
    Pattern pattern;

    @Before
    public void before() throws NoSuchFieldException, IllegalAccessException {
        Field field = ExcelExportExecutor.class.getDeclaredField("EVAL_PATTERN");
        field.setAccessible(true);
        pattern = (Pattern) field.get(null);
    }

    @Test
    public void test() {
        Matcher matcher = pattern.matcher("` 3 * ( 1 + 4 ) `");
        Assert.assertTrue(matcher.find());
        Assert.assertEquals(matcher.group(1), " 3 * ( 1 + 4 ) ");
    }

    @Test
    public void test2() {
        String value = "` 3 * ( 1 + 4 ) ` && `  1 - 1 `";
        Matcher matcher = pattern.matcher(value);
        while (matcher.find()) {
            value = value.replace(matcher.group(), matcher.group(1));
        }
        Assert.assertEquals(value, " 3 * ( 1 + 4 )  &&   1 - 1 ");
    }


    @Test
    public void test3() throws Exception {
        String[] INDEX_MAP = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
        for (int index = 0; index < 100; index++) {
            int h = index / 26;
            int l = index % 26;
            System.out.println(h <= 0 ? INDEX_MAP[l] : INDEX_MAP[h - 1] + INDEX_MAP[l]);
        }
    }

    @Test
    public void test4() {
        String expr = "sum(${column}2:${column}`${row} - 1`)";
        Matcher matcher = pattern.matcher(expr);
        while (matcher.find()) {
            System.out.println(matcher.group(1));
        }
    }

    @Test
    public void test5() throws Exception {
        ScriptEngine engine = new ScriptEngineManager().getEngineByName("js");
        System.out.println(engine.eval("'abc'"));
    }
}
