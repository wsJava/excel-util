package top.lvjp.excel.utils;

import top.lvjp.excel.constant.ReadOperatorEnum;

/**
 * @author lvjp
 * @date 2020/8/24
 */
public class ReadResult<T> {
    private T value;
    private ReadOperatorEnum curOperator;

    private ReadResult(T value, ReadOperatorEnum operator) {
        this.value = value;
        this.curOperator = operator;
    }

    public static <T> ReadResult<T> add(T value) {
        return new ReadResult<>(value, ReadOperatorEnum.CONTINUE);
    }

    public static <T> ReadResult<T> skip() {
        return new ReadResult<>(null, ReadOperatorEnum.SKIP);
    }

    public static <T> ReadResult<T> exit() {
        return new ReadResult<>(null, ReadOperatorEnum.EXIT);
    }

    public static <T> ReadResult<T> addAndExit(T value) {
        return new ReadResult<>(value, ReadOperatorEnum.ADD_EXIT);
    }

    public T get() {
        return value;
    }

    public ReadOperatorEnum curOperator() {
        return curOperator;
    }
}
