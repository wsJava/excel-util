package top.lvjp.excel.operate;

import org.apache.poi.ss.usermodel.Row;

/**
 * 读取excel的接口
 *
 * @author lvjp
 * @param <T> 生成的对象类型
 */
public interface Reader<T> {

    /**
     * @param row
     * 该方法抛出RuntimeException时程序会捕获并继续运行，若需要产生异常即停止运行请抛其他异常
     */
    T read(Row row);

}
