package com.example.demo.controller.excel;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Map;

@Controller
public class ExcelController2 {
    @Value("${testPath}")
    private String dir;


    @RequestMapping(value="/excel15", method= RequestMethod.GET)
    public ModelAndView showExcel15(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        String tempFileName = request.getParameter("temp");
        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象，参数“report”就是Excel模板中定义的单元格区域的名称
        Table table = sheet.openTableByDefinedName("report", 10, 5, true);
        //给区域中的单元格赋值
        table.getDataFields().get(0).setValue("轮胎");
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue("120");
        table.getDataFields().get(3).setValue("500");
        table.getDataFields().get(4).setValue("120%");
        //循环下一行
        table.nextRow();
        //关闭table对象
        table.close();

        //定义单元格对象，参数“year”就是Excel模板中定义的单元格的名称
        Cell cellYear = sheet.openCellByDefinedName("year");
        // 给单元格赋值
        Calendar c=new GregorianCalendar();
        int year=c.get(Calendar.YEAR);//获取年份
        cellYear.setValue(year + "年");

        Cell cellName = sheet.openCellByDefinedName("name");
        cellName.setValue("张三");

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"exl15//"+tempFileName, OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel15");
        return mv;
    }

    @RequestMapping(value="/excel16", method= RequestMethod.GET)
    public ModelAndView showExcel16(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"test16.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value="/excel17", method= RequestMethod.GET)
    public ModelAndView showExcel17(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象
        Table table = sheet.openTable("B4:F11");

        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for(int i = 1; i <= rowCount; i++){
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue( i + "月");
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue("120");
            table.getDataFields().get(3).setValue("500");
            table.getDataFields().get(4).setValue("120%");
            table.nextRow();//循环下一行，此行必须
        }

        //关闭table对象
        table.close();

        //定义Table对象
        Table table2 = sheet.openTable("B13:F16");
        int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
        for(int i = 1; i <= rowCount2; i++){
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue( i + "季度");
            table2.getDataFields().get(1).setValue("300");
            table2.getDataFields().get(2).setValue("300");
            table2.getDataFields().get(3).setValue("300");
            table2.getDataFields().get(4).setValue("100%");
            table2.nextRow();
        }

        //关闭table对象
        table2.close();

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"test16.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value="/excel18", method= RequestMethod.GET)
    public ModelAndView showExcel18(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象，参数“report1”为Excel中定义的名称，“4”为名称指定区域的行数，
        //“5”为名称指定区域的列数，“true”表示表格会按实际数据行数自动扩展
        Table table = sheet.openTableByDefinedName("report", 4, 5, true);

        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for(int i = 1; i <= rowCount; i++){
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue( i + "月");
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue("120");
            table.getDataFields().get(3).setValue("500");
            table.getDataFields().get(4).setValue("120%");
            table.nextRow();//循环下一行，此行必须
        }

        //关闭table对象
        table.close();

        //定义Table对象
        Table table2 = sheet.openTableByDefinedName("report2", 4, 5, true);
        int rowCount2 = 4;//假设将要自动填充数据的实际记录条数为12
        for(int i = 1; i <= rowCount2; i++){
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue( i + "季度");
            table2.getDataFields().get(1).setValue("300");
            table2.getDataFields().get(2).setValue("300");
            table2.getDataFields().get(3).setValue("300");
            table2.getDataFields().get(4).setValue("100%");
            table2.nextRow();
        }

        //关闭table对象
        table2.close();

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"test16.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value="/excel19", method= RequestMethod.GET)
    public ModelAndView showExcel19(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.getRibbonBar().setTabVisible("TabHome", true);//开始
        poCtrl.getRibbonBar().setTabVisible("TabFormulas", false);//公式
        poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
        poCtrl.getRibbonBar().setTabVisible("TabData", false);//数据
        poCtrl.getRibbonBar().setTabVisible("TabReview", false);//审阅
        poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图
        poCtrl.getRibbonBar().setTabVisible("TabPageLayoutExcel", false);//页面布局

        poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮

        poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板


        poCtrl.webOpen(dir+"test16.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }



    private  PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }









}
