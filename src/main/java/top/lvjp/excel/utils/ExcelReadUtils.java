package top.lvjp.excel.utils;

import org.apache.poi.ss.usermodel.Cell;

import java.math.BigDecimal;
import java.util.Date;

/**
 * @author lvjp
 * @date 2020/8/23
 */
public class ExcelReadUtils {

    /*
     * *****************************
     * 从 Excel 表格中获取数据
     * 以下方法可能抛出 NPE, Cell 可能为 null
     * *****************************
     */

    public static Boolean getBoolFromCell(Cell cell) {
        return cell.getBooleanCellValue();
    }

    public static String getStrFromCell(Cell cell) {
        return cell.getStringCellValue().trim();
    }

    public static Integer getIntFromCell(Cell cell) {
        return (int) cell.getNumericCellValue();
    }

    public static Long getLongFromCell(Cell cell) {
        return (long) cell.getNumericCellValue();
    }

    public static BigDecimal getDecimalFromCell(Cell cell) {
        return getDecimalFromCell(cell, 2);
    }

    public static BigDecimal getDecimalFromCell(Cell cell, int scale) {
        return BigDecimal.valueOf(cell.getNumericCellValue()).setScale(scale, BigDecimal.ROUND_DOWN);
    }

    public static Date getDateFromCell(Cell cell) {
        return cell.getDateCellValue();
    }


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
        return getStrFromCell(cell);
    }

    public static Integer getIntFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return getIntFromCell(cell);
    }

    public static Long getLongFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return getLongFromCell(cell);
    }

    public static BigDecimal getDecimalFromCellDefaultNull(Cell cell, int scale) {
        if (cell == null) {
            return null;
        }
        return getDecimalFromCell(cell, scale);
    }

    public static BigDecimal getDecimalFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return getDecimalFromCell(cell);
    }

    public static Date getDateFromCellDefaultNull(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.getDateCellValue();
    }
}
