package com.exl;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import java.io.File;
import java.io.IOException;


/*
 * @Author maiBangMin
 * @Description 创建Excel文件
 * @Date 11:22 下午 2020/9/21
 * @Param 
 * @return 
 **/
public class CreateExcel {

    public static void main(String[] args) throws IOException, WriteException {

        // 1.创建excel文件
        File file = new File("/Users/maibangmin/app/code/excel/Java-jxl/text7.xls");
        file.createNewFile();

        // 2.创建工作簿
        WritableWorkbook writableWorkbook = Workbook.createWorkbook(file);

        // 3.创建sheet
        WritableSheet sheet = writableWorkbook.createSheet("用户管理", 0);

        // 4.设置title
        String[] titles = {"编号","账号","密码"};

        // 5.设置单元格
        Label label = null;

        // 6.给第一行设置列名
        for (int i = 0; i < titles.length; i++) {
            /**
             * 第一个参数: 第几列
             * 第二个参数: 第几行
             * 第三个参数: 列名
             */
            label = new Label(i, 0, titles[i]);
            // 添加单元格
            sheet.addCell(label);
        }

        // 7.模拟数据库导入数据
        for (int i = 0; i < 10; i++) {
            // 添加编号,第二行第一列
            label = new Label(0, i+1,i+"");
            sheet.addCell(label);

            // 添加账号
            label = new Label(1,i+1,"10001"+i);
            sheet.addCell(label);

            // 添加密码
            label = new Label(2,i+1,"123");
            sheet.addCell(label);
        }

        // 8.写入数据
        writableWorkbook.write();

        // 9.关闭工作簿
        writableWorkbook.close();

    }

}
