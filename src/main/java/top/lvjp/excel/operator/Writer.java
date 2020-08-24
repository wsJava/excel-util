package top.lvjp.excel.operator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.Field;
import java.util.Objects;

/**
 * 写入Excel的接口
 *
 * @author lvjp
 * @date 2020/6/29
 */
public interface Writer<T> {

    /**
     * 获取表格头部字段，即每列的列名，顺序应该和 write 中的顺序保持一致，默认没有列名。需要列名请重写该方法
     *
     * @return 列名数组，如果返回 {@code null} 则不会添加列名
     */
    default String[] getHeaders() {
        return null;
    }

    /**
     * 将对象写入Excel表格的一行中
     *
     * @param row Excel 行对象
     * @param t   需要写入的对象
     */
    void write(Row row, T t);

    /**
     * 提供的一个默认的 Excel writer，将对象所有的定义字段（不包括继承的）写入表格，以字段名为列名，表格内的值为 String类型
     *
     * @param tClass 数据的类对象
     * @param <T>    数据类型
     * @return writer
     */
    static <T> Writer<T> defaultWriter(Class<T> tClass) {
        Objects.requireNonNull(tClass, "tClass must not be null");
        final Field[] fields = tClass.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
        }

        return new Writer<T>() {
            @Override
            public String[] getHeaders() {
                String[] headers = new String[fields.length];
                for (int i = 0; i < fields.length; i++) {
                    headers[i] = fields[i].getName();
                }
                return headers;
            }

            @Override
            public void write(Row row, T t) {
                String[] headers = getHeaders();
                for (int i = 0; i < headers.length; i++) {
                    Cell cell = row.createCell(i, CellType.STRING);
                    try {
                        cell.setCellValue(String.valueOf(fields[i].get(t)));
                    } catch (IllegalAccessException ignored) {
                    }
                }
            }
        };
    }
}
