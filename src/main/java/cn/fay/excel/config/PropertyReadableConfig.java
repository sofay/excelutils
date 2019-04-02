package cn.fay.excel.config;

import cn.fay.excel.util.PropertyLoader;

import java.util.HashMap;
import java.util.Map;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-28 19:47.
 */
public class PropertyReadableConfig implements ReadableConfig {
    private Map<String, String> properties = new HashMap<>();

    public PropertyReadableConfig(String propFile) {
        PropertyLoader.loader(propFile, new PropertyLoader.PropTravel() {
            @Override
            public void travel(String k, String v) {
                properties.put(k, v);
            }
        });
    }

    @Override
    public String get(String key) {
        return properties.get(key);
    }
}
