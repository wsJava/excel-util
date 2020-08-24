package top.lvjp.excel.constant;

/**
 * @author lvjp
 * @date 2020/8/24
 */
public enum ReadOperatorEnum {
    /**
     * 继续读取 excel
     */
    CONTINUE,

    /**
     * 直接退出
     */
    EXIT,

    /**
     * 添加数据并退出
     */
    ADD_EXIT,

    /**
     * 跳过当前行
     */
    SKIP;
}
