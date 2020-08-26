package top.lvjp.excel.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.math.BigDecimal;
import java.text.DateFormat;
import java.util.Date;

/**
 * @author lvjp
 * @date 2020/8/23
 */
public class ExcelReadUtils {

    /*
     * *****************************
     * 从 Excel 表格中获取数据
     * 默认值为 null
     * *****************************
     */

    public static Boolean getBoolFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.getBooleanCellValue();
    }

    public static String getStrFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.getStringCellValue();
    }

    public static String getTrimStrFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.getStringCellValue().trim();
    }

    public static Integer getIntFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                return (int) cell.getNumericCellValue();
            case STRING:
                return Integer.valueOf(cell.getStringCellValue().trim());
            default:
                return null;
        }
    }

    public static Long getLongFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                return (long) cell.getNumericCellValue();
            case STRING:
                return Long.valueOf(cell.getStringCellValue().trim());
            default:
                return null;
        }
    }

    public static Double getDoubleFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return Double.valueOf(cell.getStringCellValue().trim());
            default:
                return null;
        }
    }
}
