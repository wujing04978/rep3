package com.wj.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.jupiter.api.DynamicTest;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    @Test
    public void fun2() throws IOException {
		int x1=10;

        //如果d盘中有这个文件就要先删除这个文件

        //创建工作簿(可以理解为Excel文件)
        HSSFWorkbook workBook = new HSSFWorkbook();

        ////创建工作表  工作表（sheet）的名字叫hello
        HSSFSheet sheet = workBook.createSheet("hello");

        //第一行
        HSSFRow row01=sheet.createRow(0);//下标为0代表第一行
        HSSFCell cell0101=row01.createCell(0);//代表当前是第一行第一列
        cell0101.setCellValue("姓名");
        HSSFCell cell0102=row01.createCell(1);//代表的是第一行的第二列
        cell0102.setCellValue("年龄");


        for(int i=1;i<=5;i++){
            HSSFRow rowx=sheet.createRow(i); //i从2开始
            HSSFCell cell01=rowx.createCell(0);
            cell01.setCellValue("张三"+i);
            HSSFCell cell02=rowx.createCell(1);
            cell02.setCellValue(i+18);
        }

        workBook.write(new FileOutputStream(new File("d:\\demoadmin.xls")));
        workBook.close();

        ////////////////////////////////////////////////////////////////
        //下载操作

    }


    @Test
    public void fun1() throws IOException {
        //创建工作簿(可以理解为Excel文件)
        HSSFWorkbook workBook = new HSSFWorkbook();

        ////创建工作表  工作表（sheet）的名字叫hello
        HSSFSheet sheet = workBook.createSheet("hello");

        //第一行
        HSSFRow row01=sheet.createRow(0);//下标为0代表第一行
        HSSFCell cell0101=row01.createCell(0);//代表当前是第一行第一列
        cell0101.setCellValue("姓名");
        HSSFCell cell0102=row01.createCell(1);//代表的是第一行的第二列
        cell0102.setCellValue("年龄");


        //第二行
        HSSFRow row02=sheet.createRow(1);
        HSSFCell cell0201=row02.createCell(0);
        cell0201.setCellValue("张三");
        HSSFCell cell0202=row02.createCell(1);
        cell0202.setCellValue("18");


        //第二行
        HSSFRow row03=sheet.createRow(2);
        HSSFCell cell0301=row03.createCell(0);
        cell0301.setCellValue("张三");
        HSSFCell cell0302=row03.createCell(1);
        cell0302.setCellValue("18");

        workBook.write(new FileOutputStream(new File("d:\\demo01.xls")));
        workBook.close();
    }


}
