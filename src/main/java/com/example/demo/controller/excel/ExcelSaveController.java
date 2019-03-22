package com.example.demo.controller.excel;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.excelreader.Sheet;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelreader.Workbook;
import com.zhuozhengsoft.pageoffice.wordreader.DataRegion;
import com.zhuozhengsoft.pageoffice.wordreader.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class ExcelSaveController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping("/save/exl2")
    public String saveExl2(HttpServletRequest request, HttpServletResponse response) {
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("Info");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                //out.print(table.getDataFields().get(2).getText()+"      mmmmmmmmmmmmm          "+table.getDataFields().get(1).getText());
                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length() == 0
                ) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        System.out.println(content);
        workBook.showPage(500, 400);
        workBook.close();
        return "/resp";

    }


    @RequestMapping("/save/exl/data1")
    public String saveExlData1(HttpServletRequest request, HttpServletResponse response) {
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("B4:F13");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length() == 0
                ) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "</br>";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        workBook.showPage(500, 400);
        workBook.close();
        System.out.println(content);
        request.setAttribute("content", content);
        return "/resp";

    }

    @RequestMapping("/save/exl/data2")
    public void saveExlData2(HttpServletRequest request, HttpServletResponse response) {
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("B4:D8");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {

            //获取提交的数值
            //DataFields.Count标识的是table的列数
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称：" + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量：" + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量：" + table.getDataFields().get(2).getText();

                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        System.out.println(content);


        workBook.close();

    }

    @RequestMapping("/save/exl/data3")
    public void saveExlData3(HttpServletRequest request, HttpServletResponse response) {
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");

        String content = "";
        content += "testA1：" + sheet.openCellByDefinedName("testA1").getValue() + "<br/>";
        content += "testB1：" + sheet.openCellByDefinedName("testB1").getValue() + "<br/>";

        workBook.showPage(500, 400);
        workBook.close();
        System.out.println(content);

    }

    @RequestMapping("/save/exl/data4")
    public void saveExlData4(HttpServletRequest request, HttpServletResponse response) {
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");

        Table table = sheet.openTableByDefinedName("report");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                //out.print(table.getDataFields().get(2).getText()+"      mmmmmmmmmmmmm          "+table.getDataFields().get(1).getText());

                int colCount = table.getDataFields().size();//获取列数

                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length() == 0) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df = (DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f * 100) + "%";
                }
                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        workBook.showPage(500, 400);
        workBook.close();
        System.out.println(content);

    }


}
