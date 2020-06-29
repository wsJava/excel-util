package top.lvjp.excel.constant;

/**
 * @author lvjp
 * @date 2020/6/29
 */
public enum FileTypeEnum {

    xls("XLS"),
    xlsx("XLSX");

    private String suffix;

    FileTypeEnum(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix() {
        return suffix;
    }
}
