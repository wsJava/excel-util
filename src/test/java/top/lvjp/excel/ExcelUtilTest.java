package top.lvjp.excel;

import org.apache.poi.ss.usermodel.Row;
import top.lvjp.excel.operator.Writer;
import top.lvjp.excel.utils.ExcelReadUtils;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtilTest {

    public static void main(String[] args) throws IOException {
//        ExcelUtil.addDataToExcelFile(getStudent(), new File("student.xlsx"), 2, 9, Writer.defaultWriter(Student.class));
        File file = ExcelUtil.writeNewExcel(getStudent(), "student.xlsx", Writer.defaultWriter(Student.class));

        List<Student> students = ExcelUtil.readExcel(file, 0, 1, ExcelUtilTest::readStudent);

        System.out.println(students);
    }

    private static ReadResult<Student> readStudent(Row row) {
        if (row == null) {
            return ReadResult.skip();
        }
        // 读取到第6行就退出
        if (row.getRowNum() > 5) {
            return ReadResult.exit();
        }

        // 如果id这个 cell（格子）为null，就说明excel格式不正确,捕获异常，添加相关信息，继续抛出
        Integer id;
        try {
             id = ExcelReadUtils.getIntFromCellDefaultNull(row.getCell(0));
        } catch (Exception e) {
            throw new RuntimeException("excel 读取异常，当前行号："+row.getRowNum(), e);
        }
        // 姓名这格不存在就默认null
        String s = ExcelReadUtils.getStrFromCellDefaultNull(row.getCell(1));
        return ReadResult.add(new Student(id, s));
    }

    private static List<Student> getStudent() {
        List<Student> list = new ArrayList<>(100);
        for (int i = 0; i < 10; i++) {
            list.add(new Student(i, "stu-"+i));
        }
        return list;
    }


    static class Student{
        Integer id;
        String name;
        public Student(Integer id, String name) {
            this.id = id;
            this.name = name;
        }

        public Integer getId() {
            return id;
        }

        public void setId(Integer id) {
            this.id = id;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        @Override
        public String toString() {
            return "Student{" +
                    "id=" + id +
                    ", name='" + name + '\'' +
                    '}';
        }
    }
}