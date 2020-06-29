package top.lvjp.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import top.lvjp.excel.constant.FileTypeEnum;
import top.lvjp.excel.operate.Reader;
import top.lvjp.excel.operate.Writer;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * @author lvjp
 * @date 2020/6/29
 */
public class ExcelUtil {

    private static final Logger log = LoggerFactory.getLogger(ExcelUtil.class);


    public <T> List<T> readExcel(File file, int sheetIndex, int startRow, Reader<T> reader) throws IOException {
        Workbook workbook = getWorkbook(file);
        return readExcel(workbook, sheetIndex, startRow, reader);
    }

    public <T> List<T> readExcel(InputStream inputStream, FileTypeEnum fileType, int sheetIndex, int startRow, Reader<T> reader) throws IOException {
        Workbook workbook;
        if (FileTypeEnum.xls == fileType) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return readExcel(workbook, 0, 1, reader);
    }

    public <T> List<T> readExcel(Workbook workbook, int sheetIndex, int startRow, Reader<T> reader) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        List<T> list = new ArrayList<>(sheet.getLastRowNum());

        for (int rowIndex = startRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            try {
                list.add(reader.read(row));
            } catch (RuntimeException e) {
                log.error("读取Excel行失败，当前行号：{}", rowIndex, e);
            }
        }
        return list;
    }

    public Workbook getWorkbook(File file) throws IOException {
        Workbook workbook;
        String fileName = file.getName();
        int index = fileName.lastIndexOf(".");
        String fileType = fileName.substring(index + 1).toUpperCase();

        try (BufferedInputStream buffer = new BufferedInputStream(new FileInputStream(file))) {
            if (FileTypeEnum.xls.getSuffix().equals(fileType)) {
                workbook = new HSSFWorkbook(buffer);
            } else if (FileTypeEnum.xlsx.getSuffix().equals(fileType)) {
                workbook = new XSSFWorkbook(buffer);
            } else {
                throw new RuntimeException("文件格式不正确");
            }
        }
        return workbook;
    }


    public <T> void writeExcel(List<T> data, OutputStream out, int startRow, Writer<T> writer) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        writeExcel(data, workbook, startRow, writer);
        workbook.write(out);
        workbook.close();
    }


    private <T> void writeExcel(List<T> data, Workbook workbook, int startRow, Writer<T> writer) {

        Sheet sheet = workbook.createSheet();
        String[] headers = writer.getHeaders();

        Row header = sheet.createRow(startRow++);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
        }

        for (T t : data) {
            Row row = sheet.createRow(startRow++);
            writer.write(row, t);
        }
    }
}
