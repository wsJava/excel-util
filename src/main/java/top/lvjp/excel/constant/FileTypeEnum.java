package top.lvjp.excel.constant;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author lvjp
 * @date 2020/6/29
 */
public enum FileTypeEnum {

    /**
     * excel 97-2003 版， 对应 {@link org.apache.poi.hssf.usermodel.HSSFWorkbook}
     */
    XLS("XLS"),

    /**
     * excel 20007-2016 的文件格式， 对应 {@link org.apache.poi.xssf.usermodel.XSSFWorkbook}
     */
    XLSX("XLSX");

    private String suffix;

    FileTypeEnum(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix() {
        return suffix;
    }

    /**
     * 根据文件名获取 Excel 文件类型
     *
     * @param fileName 文件名
     * @return FileTypeEnum
     */
    public static FileTypeEnum ofFileName(String fileName) {
        if (fileName == null || "".equals(fileName)) {
            throw new IllegalArgumentException("请输入正确文件名");
        }
        int index = fileName.lastIndexOf(".");
        if (index < 0) {
            throw new IllegalArgumentException("文件名不合法，fileName=" + fileName);
        }
        String fileType = fileName.substring(index + 1).toUpperCase();
        for (FileTypeEnum value : FileTypeEnum.values()) {
            if (value.suffix.equals(fileType)) {
                return value;
            }
        }
        throw new IllegalArgumentException("不支持的文件类型，fileType=" + fileType);
    }
}
