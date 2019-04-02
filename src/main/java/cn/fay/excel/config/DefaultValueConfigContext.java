package cn.fay.excel.config;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-28 19:55.
 */
public class DefaultValueConfigContext {
    private final static String DEFAULT_VALUE_CONFIG_FILE = "default_value.properties";
    private static ReadableConfig config;
    static {
        config = new PropertyReadableConfig(DEFAULT_VALUE_CONFIG_FILE);
    }

    public static String transDefaultValue(String key) {
        return config.get(key);
    }
}
