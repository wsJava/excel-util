package top.lvjp.excel.operate;

import org.apache.poi.ss.usermodel.Row;

/**
 * 写入Excel的接口
 *
 * @author lvjp
 * @date 2020/6/29
 */
public interface Writer<T> {

    String[] getHeaders();

    void write(Row row, T t);
}
