package top.lvjp.excel.operator;

import org.apache.poi.ss.usermodel.Row;
import top.lvjp.excel.ReadResult;

/**
 * 读取excel的接口
 *
 * @param <T> 生成的对象类型
 * @author lvjp
 */
public interface Reader<T> {

    /**
     * 读取 excel 行，注意如遇到空行 row 可能为 null, 请自行处理。
     * 另外请自行处理方法中抛出的异常，并决定后续操作；若该方法抛出异常，不做任何处理直接向上抛出。
     *
     * @param row excel 的行对象，可能为 null
     * @return ReadResult 持有读取行得到的数据对象，并决定下一步操作，退出或者继续读取
     */
    ReadResult<T> read(Row row);

}
