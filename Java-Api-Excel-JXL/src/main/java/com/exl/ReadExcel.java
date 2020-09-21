package com.exl;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;


/*
 * @Author maiBangMin
 * @Description 读取Excel文件
 * @Date 11:22 下午 2020/9/21
 * @Param 
 * @return 
 **/
public class ReadExcel {

    public static void main(String[] args) throws IOException, WriteException {

        // 1.创建工作簿
        File file = new File("/Users/maibangmin/app/code/excel/Java-jxl/text7.xls");
        WritableWorkbook workbook = Workbook.createWorkbook(file);

        // 2.获取第一个工作表
        WritableSheet[] sheets = workbook.getSheets();

        // 3.获取数据
        for (Sheet sheet:sheets) {
            // 3.1获取所有的行数
            System.out.println("行数:" + sheet.getRows());
            // 3.2获取所有的列数
            System.out.println("列数" + sheet.getColumns());
            // 3.3获取单元格内的数据
            for (int i = 0; i < sheet.getRows(); i++) {
                for (int j = 0; j < sheet.getColumns(); j++) {
                    String contents = sheet.getCell(j, i).getContents();
                    System.out.println(contents);
                }
            }
        }
        System.out.println();
        // 4.关闭资源
        workbook.close();
    }

}
