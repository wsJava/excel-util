package top.lvjp.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.lvjp.excel.constant.FileTypeEnum;
import top.lvjp.excel.constant.ReadOperatorEnum;
import top.lvjp.excel.operator.Reader;
import top.lvjp.excel.operator.Writer;

import java.io.*;
import java.util.*;

/**
 * Excel 读写工具，提供了一些列读取或写入 excel 的方法，只需要传入读写的策略实例，和指定的读写范围，就可以得到相应的数据或excel。
 * 注意：行是从 0 开始计算
 *
 * @author lvjp
 * @date 2020/6/29
 */
public class ExcelUtil {

    /**
     * 读取 excel 文件
     *
     * @param file       excel 文件
     * @param sheetIndex 指定 sheet 页
     * @param startRow   开始行 0
     * @param reader     读取excel的对象
     * @param <T>        数据类型
     * @return 读取 excel 得到的数据
     * @throws IOException 生成工作薄对象时可能抛出 IO 异常
     */
    public static <T> List<T> readExcel(File file, int sheetIndex, int startRow, Reader<T> reader) throws IOException {
        Objects.requireNonNull(file, "file must not be null");
        Workbook workbook = getWorkbook(file);
        return readExcel(workbook, sheetIndex, startRow, reader);
    }

    /**
     * 读取输入流中的 excel 表格
     *
     * @param inputStream 输入流
     * @param fileType    指定 excel 类型
     * @param sheetIndex  指定的 sheet 页
     * @param startRow    开始行， 0 开始
     * @param reader      读取excel的对象
     * @param <T>         数据类型
     * @return 读取 excel 得到的数据
     * @throws IOException 生成工作薄对象时可能抛出 IO 异常
     */
    public static <T> List<T> readExcel(InputStream inputStream, FileTypeEnum fileType, int sheetIndex, int startRow, Reader<T> reader) throws IOException {
        Objects.requireNonNull(inputStream, "inputStream must not be null");
        Objects.requireNonNull(fileType, "fileType must not be null");
        Workbook workbook;
        if (FileTypeEnum.XLS == fileType) {
            workbook = new HSSFWorkbook(inputStream);
        } else if (FileTypeEnum.XLSX == fileType) {
            workbook = new XSSFWorkbook(inputStream);
        } else {
            throw new RuntimeException("不支持的文件类型 fileType=" + fileType);
        }
        return readExcel(workbook, sheetIndex, startRow, reader);
    }

    /**
     * 读取 Excel 工作薄的数据
     *
     * @param workbook   工作薄
     * @param sheetIndex 工作薄指定的sheet 页
     * @param startRow   开始读取的行，注意是从 0 开始
     * @param reader     读取excel的对象
     * @param <T>        数据类型
     * @return 读取 excel 得到的数据
     */
    public static <T> List<T> readExcel(Workbook workbook, int sheetIndex, int startRow, Reader<T> reader) {
        Objects.requireNonNull(workbook, "workbook must not be null");
        Objects.requireNonNull(reader, "reader must not be null");
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        List<T> list = new ArrayList<>(sheet.getLastRowNum());

        for (int rowIndex = startRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            ReadResult<T> result = reader.read(row);

            ReadOperatorEnum curOperator = result.curOperator();
            switch (curOperator) {
                case CONTINUE:
                    list.add(result.get());
                    break;
                case SKIP:
                    break;
                case ADD_EXIT:
                    list.add(result.get());
                case EXIT:
                    return list;
            }

        }
        return list;
    }

    /**
     * 将数据写入到指定 Excel 文件中，如果文件已存在则覆盖原有内容，不存在则创建
     *
     * @param data     待写入的数据
     * @param fileName 写入的文件名，如果不存在则主动创建
     * @param writer   具体写入的方式策略
     * @param <T>      数据类型
     * @return file excel 文件
     * @throws IOException IO异常时抛出，需自行处理
     */
    public static <T> File writeNewExcel(List<T> data, String fileName, Writer<T> writer) throws IOException {
        Objects.requireNonNull(fileName, "fileName must not be null");
        File file = new File(fileName);
        if (!file.exists()) {
            boolean success = file.createNewFile();
            if (!success) {
                throw new RuntimeException("create file=" + fileName + "失败");
            }
        }
        FileTypeEnum fileType = FileTypeEnum.ofFileName(fileName);
        try (OutputStream out = new FileOutputStream(file)) {
            writeExcelToOutStream(data, out, fileType, writer);
        }
        return file;
    }

    /**
     * 将数据追加到已存在的excel文件中
     *
     * @param data       需要写入的数据
     * @param file       已存在的excel文件
     * @param sheetIndex 指定的 sheet 页， 0开始
     * @param startRow   指定的开始行，0开始
     * @param writer     具体写入的方式策略
     * @param <T>        数据类型
     */
    public static <T> void addDataToExcelFile(List<T> data, File file, int sheetIndex, int startRow, Writer<T> writer) throws IOException {
        Objects.requireNonNull(file, "file must not be null");
        try (Workbook workbook = getWorkbook(file)) {
            writeExcelToWorkbook(data, workbook, sheetIndex, startRow, writer);
            try (OutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }
        }
    }

    /**
     * 将数据写入到 输出流 中
     */
    public static <T> void writeExcelToOutStream(List<T> data, OutputStream out, FileTypeEnum fileType, Writer<T> writer) throws IOException {
        Objects.requireNonNull(out, "out must not be null");
        Objects.requireNonNull(writer, "writer must not be null");
        try (Workbook workbook = createNewWorkbook(fileType)) {
            writeExcelToWorkbook(data, workbook, 0, 0, writer);
            workbook.write(out);
        }
    }

    /**
     * 将数据写入 excel 的工作薄
     *
     * @param data       待写入的数据
     * @param workbook   工作薄
     * @param sheetIndex 写入的sheet页
     * @param startRow   开始行，从0开始
     * @param writer     具体写入的方式策略
     * @param <T>        数据类型
     */
    public static <T> void writeExcelToWorkbook(List<T> data, Workbook workbook, int sheetIndex, int startRow, Writer<T> writer) {
        Objects.requireNonNull(data, "data must not be null");
        Objects.requireNonNull(workbook, "workbook must not be null");
        Objects.requireNonNull(writer, "writer must not be null");

        Sheet sheet = getOrCreateSheet(workbook, sheetIndex);
        String[] headers = writer.getHeaders();

        if (headers != null) {
            Row header = sheet.createRow(startRow++);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = header.createCell(i);
                cell.setCellValue(headers[i]);
            }
        }

        for (T t : data) {
            Row row = sheet.createRow(startRow++);
            writer.write(row, t);
        }
    }

    private static Sheet getOrCreateSheet(Workbook workbook, int sheetIndex) {
        Sheet sheet;
        try {
            sheet = workbook.getSheetAt(sheetIndex);
        } catch (IllegalArgumentException e) {
            sheet = workbook.createSheet();
        }
        return sheet;
    }

    /**
     * 根据 excel 文件，获取对应的工作薄对象
     */
    public static Workbook getWorkbook(File file) throws IOException {
        Objects.requireNonNull(file, "file must not be null");
        Workbook workbook;
        String fileName = file.getName();
        FileTypeEnum fileType = FileTypeEnum.ofFileName(fileName);

        try (BufferedInputStream buffer = new BufferedInputStream(new FileInputStream(file))) {
            if (FileTypeEnum.XLS == fileType) {
                workbook = new HSSFWorkbook(buffer);
            } else if (FileTypeEnum.XLSX == fileType) {
                workbook = new XSSFWorkbook(buffer);
            } else {
                throw new RuntimeException("不支持的文件类型");
            }
        }
        return workbook;
    }


    /**
     * 创建一个指定excel类型的工作薄
     */
    public static Workbook createNewWorkbook(FileTypeEnum fileTypeEnum) {
        Workbook workbook;
        if (FileTypeEnum.XLS == fileTypeEnum) {
            workbook = new HSSFWorkbook();
        } else if (FileTypeEnum.XLSX == fileTypeEnum) {
            workbook = new XSSFWorkbook();
        } else {
            throw new RuntimeException("文件类型不支持");
        }
        return workbook;
    }
}
