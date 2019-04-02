package cn.fay.excel.util;


import java.io.IOException;
import java.net.URL;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 * @author fay  fay9395@gmail.com
 * @date 2019-03-28 19:20.
 */
public class PropertyLoader {

    public static Map<String, String> loader(String fileName, boolean failOnSameKey) {
        Map<String, String> result = new HashMap<>();
        loader(fileName, (key, value) -> {
            if (result.containsKey(key) && failOnSameKey) {
                throw new RuntimeException("find same property key: " + key);
            }
        });
        return result;
    }

    public static void loader(String fileName, PropTravel propTravel) {
        try {
            Enumeration<URL> urls = ClassLoader.getSystemResources(fileName);
            if (urls != null) {
                while (urls.hasMoreElements()) {
                    URL url = urls.nextElement();
                    Properties prop = new Properties();
                    prop.load(url.openStream());
                    for (Object key : prop.keySet()) {
                        String eval = prop.getProperty(key.toString());
                        propTravel.travel(key.toString(), eval);
                    }
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    public interface PropTravel {
        void travel(String key, String value);
    }
}
