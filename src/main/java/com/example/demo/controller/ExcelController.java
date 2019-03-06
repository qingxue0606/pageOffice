package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@Controller
public class ExcelController {
    @RequestMapping(value="/excel", method= RequestMethod.GET)
    public ModelAndView showExcel(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1); //添加自定义按钮

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Cell对象
        Cell cellB4 = sheet.openCell("B4");
        //给单元格赋值
        cellB4.setValue("1月");

        Cell cellC4 = sheet.openCell("C4");
        cellC4.setValue("300");

        Cell cellD4 = sheet.openCell("D4");
        cellD4.setValue("270");

        Cell cellE4 = sheet.openCell("E4");
        cellE4.setValue("270");

        Cell cellF4 = sheet.openCell("F4");
        DecimalFormat df=(DecimalFormat) NumberFormat.getInstance();
        cellF4.setValue(df.format( 270.00 / 300*100)+"%");

        poCtrl.setWriter(workBook);


        poCtrl.setSaveFilePage("/save");//设置保存的action
        poCtrl.addCustomToolButton("盖章","AddSeal",2); //添加自定义盖章按钮
        poCtrl.webOpen("d:\\test\\test.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }
    @RequestMapping(value="/excel2", method= RequestMethod.GET)
    public ModelAndView showExcel2(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存","Save",1); //添加自定义按钮
//定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义table对象，设置table对象的设置范围
        Table table = sheet.openTable("B4:F13");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");

        poCtrl.setWriter(workBook);


        poCtrl.setSaveDataPage("/save/exl2");

        poCtrl.addCustomToolButton("盖章","AddSeal",2); //添加自定义盖章按钮
        poCtrl.webOpen("d:\\test\\test2.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }


    @RequestMapping(value="/excel3", method= RequestMethod.GET)
    public ModelAndView showExcel3(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.setCaption("使用OpenTable给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Table对象
        Table table = sheet.openTable("B4:F13");
        for(int i=0; i < 50; i++)
        {
            table.getDataFields().get(0).setValue("产品 " + i);
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue(String.valueOf(100+i));
            table.nextRow();
        }
        table.close();

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        poCtrl.webOpen("d:\\test\\test3.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }

    @RequestMapping(value="/excel4", method= RequestMethod.GET)
    public ModelAndView showExcel4(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("导入文件", "importData()", 16);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        poCtrl.setWriter(wb);
        poCtrl.setSaveDataPage("/save/exl/data1");

        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("exce1");
        return mv;
    }


    @RequestMapping(value="/excel5", method= RequestMethod.GET)
    public ModelAndView showExcel5(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        Workbook  workBoook=new Workbook();
        workBoook.setDisableSheetRightClick(true);//禁止当前工作表鼠标右键
//workBoook.setDisableSheetDoubleClick(true);//禁止当前工作表鼠标双击
//workBoook.setDisableSheetSelection(true);//禁止在当前工作表中选择内容
        poCtrl.setWriter(workBoook);

        poCtrl.webOpen("d:\\test\\test3.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }








}
