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
import java.util.Map;

@Controller
public class ExcelController {
    @Value("${testPath}")
    private String dir;
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
        poCtrl.webOpen(dir+"test2.xls", OpenModeType.xlsNormalEdit,"张三");
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
        poCtrl.webOpen(dir+"test2.xls", OpenModeType.xlsSubmitForm,"张三");
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
        poCtrl.webOpen(dir+"test3.xls", OpenModeType.xlsSubmitForm,"张三");
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

        poCtrl.webOpen(dir+"test3.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }

    @RequestMapping(value="/excel6", method= RequestMethod.GET)
    public ModelAndView showExcel6(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        //合并单元格
        sheet.openTable("B2:F2").merge();
        Cell cB2 = sheet.openCell("B2");
        cB2.setValue("北京某公司产品销售情况");
        //设置水平对齐方式
        cB2.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        cB2.setForeColor(Color.red);
        cB2.getFont().setSize(16);

        sheet.openTable("B4:B6").merge();
        Cell cB4 = sheet.openCell("B4");
        cB4.setValue("A产品");
        //设置水平对齐方式
        cB4.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        //设置垂直对齐方式
        cB4.setVerticalAlignment(XlVAlign.xlVAlignCenter);

        sheet.openTable("B7:B9").merge();
        Cell cB7 = sheet.openCell("B7");
        cB7.setValue("B产品");
        cB7.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        cB7.setVerticalAlignment(XlVAlign.xlVAlignCenter);

        poCtrl.setWriter(wb);
        poCtrl.setCustomToolbar(false);

        poCtrl.webOpen(dir+"test6.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel6");
        return mv;
    }


    @RequestMapping(value="/excel7", method= RequestMethod.GET)
    public ModelAndView showExcel7(HttpServletRequest request, Map<String,Object> map){

        ModelAndView mv = new ModelAndView("excel/excel7");
        return mv;
    }
    @RequestMapping(value="/excel8", method= RequestMethod.GET)
    public ModelAndView showExcel8(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        String userName=request.getParameter("userName");


        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        Table tableA = sheet.openTable("C4:D6");
        Table tableB = sheet.openTable("C7:D9");

        tableA.setSubmitName("tableA");
        tableB.setSubmitName("tableB");


        //根据登录用户名设置数据区域可编辑性
        String strInfo = "";

        //A部门经理登录后
        if (userName.equals("zhangsan"))
        {
            strInfo = "A部门经理，所以只能编辑A部门的产品数据";
            tableA.setReadOnly(false);
            tableB.setReadOnly(true);
        }
        //B部门经理登录后
        else
        {
            strInfo = "B部门经理，所以只能编辑B部门的产品数据";
            tableA.setReadOnly(true);
            tableB.setReadOnly(false);
        }

        poCtrl.setWriter(wb);
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setSaveFilePage("SaveFile.jsp");


        poCtrl.webOpen(dir+"test8.xls", OpenModeType.xlsSubmitForm,userName);
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel8");
        return mv;
    }
    @RequestMapping(value="/excel9", method= RequestMethod.GET)
    public ModelAndView showExcel9(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");
        // 设置背景
        Table backGroundTable = sheet.openTable("A1:P200");
        //设置表格边框样式
        backGroundTable.getBorder().setLineColor(Color.white);

        // 设置单元格边框样式
        Border C4Border = sheet.openTable("C4:C4").getBorder();
        C4Border.setWeight(XlBorderWeight.xlThick);
        C4Border.setLineColor(Color.yellow);
        C4Border.setBorderType(XlBorderType.xlAllEdges);

        // 设置单元格边框样式
        Border B6Border = sheet.openTable("B6:B6").getBorder();
        B6Border.setWeight(XlBorderWeight.xlHairline);
        B6Border.setLineColor(Color.magenta);
        B6Border.setLineStyle(XlBorderLineStyle.xlSlantDashDot);
        B6Border.setBorderType(XlBorderType.xlAllEdges);

        //设置表格边框样式
        Table titleTable = sheet.openTable("B4:F5");
        titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
        titleTable.getBorder().setLineColor(new Color(0, 128, 128));
        titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

        //设置表格边框样式
        Table bodyTable2 = sheet.openTable("B6:F15");
        bodyTable2.getBorder().setWeight(XlBorderWeight.xlThick);
        bodyTable2.getBorder().setLineColor(new Color(0, 128, 128));
        bodyTable2.getBorder().setBorderType(XlBorderType.xlAllEdges);

        poCtrl.setWriter(wb);



        poCtrl.webOpen(dir+"test9.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel9");
        return mv;
    }

    @RequestMapping(value="/excel10", method= RequestMethod.GET)
    public ModelAndView showExcel10(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        Workbook wb = new Workbook();
        Sheet sheet = wb.openSheet("Sheet1");

        Cell cC3 = sheet.openCell("C3");
        //设置单元格背景样式
        cC3.setBackColor( Color.LIGHT_GRAY);
        cC3.setValue( "一月");
        cC3.setForeColor(Color.white);
        cC3.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

        Cell cD3 = sheet.openCell("D3");
        //设置单元格背景样式
        cD3.setBackColor( Color.lightGray);
        cD3.setValue( "二月");
        cD3.setForeColor(Color.white);
        cD3.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        Cell cE3 = sheet.openCell("E3");
        //设置单元格背景样式
        cE3.setBackColor( Color.lightGray);
        cE3.setValue( "三月");
        cE3.setForeColor(Color.white);
        cE3.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        Cell cB4 = sheet.openCell("B4");
        //设置单元格背景样式
        cB4.setBackColor( new Color(10,254,254));
        cB4.setValue( "住房");
        cB4.setForeColor( new Color(10,150,150));
        cB4.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        Cell cB5 = sheet.openCell("B5");
        //设置单元格背景样式
        cB5.setBackColor( new Color(10,150,150));
        cB5.setValue( "三餐");
        cB5.setForeColor( new Color(10,100,250));
        cB5.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        Cell cB6 = sheet.openCell("B6");
        //设置单元格背景样式
        cB6.setBackColor(new Color(200,200,100) );
        cB6.setValue( "车费");
        cB6.setForeColor( new Color(10,150,150));
        cB6.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        Cell cB7 = sheet.openCell("B7");
        //设置单元格背景样式
        cB7.setBackColor( new Color(80,50,80));
        cB7.setValue( "通讯");
        cB7.setForeColor( new Color(10,150,150));
        cB7.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

        //绘制表格线
        Table titleTable = sheet.openTable("B3:E10");
        titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
        titleTable.getBorder().setLineColor(new Color(0, 128, 128));
        titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

        sheet.openTable("B1:E2").merge();//合并单元格
        sheet.openTable("B1:E2").setRowHeight( 30);//设置行高
        Cell B1 = sheet.openCell("B1");
        //设置单元格文本样式
        B1.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        B1.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        B1.setForeColor( new Color(0,128,128));
        B1.setValue( "出差开支预算");
        B1.getFont().setBold(true);
        B1.getFont().setSize(25);

        poCtrl.setWriter(wb);



        poCtrl.webOpen(dir+"test10.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel9");
        return mv;
    }
    @RequestMapping(value="/excel11", method= RequestMethod.GET)
    public ModelAndView showExcel11(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义table对象，设置table对象的设置范围
        Table table = sheet.openTable("B4:D8");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");

        // 设置响应单元格点击事件的js function
        poCtrl.setJsFunction_OnExcelCellClick("OnCellClick()");

        poCtrl.setWriter(workBook);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存页面
        poCtrl.setSaveDataPage("/save/exl/data2");



        poCtrl.webOpen(dir+"test11.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel11");
        return mv;
    }


    @RequestMapping(value="/excel12", method= RequestMethod.GET)
    public ModelAndView showExcel12(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        //定义Cell对象
        Cell cellB4 = sheet.openCell("B4");
        //给单元格赋值
        cellB4.setValue("1月");
        //设置字体颜色
        cellB4.setForeColor(Color.red);

        Cell cellC4 = sheet.openCell("C4");
        cellC4.setValue("300");
        cellC4.setForeColor(Color.blue);

        Cell cellD4 = sheet.openCell("D4");
        cellD4.setValue("270");
        cellD4.setForeColor(Color.orange);

        Cell cellE4 = sheet.openCell("E4");
        cellE4.setValue("270");
        cellE4.setForeColor(Color.green);

        Cell cellF4 = sheet.openCell("F4");
        DecimalFormat df=(DecimalFormat)NumberFormat.getInstance();
        cellF4.setValue(df.format( 270.00 / 300*100)+"%");
        cellF4.setForeColor(Color.gray);

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文件



        poCtrl.webOpen(dir+"test12.xls", OpenModeType.xlsSubmitForm,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel8");
        return mv;
    }

    @RequestMapping(value="/excel13", method= RequestMethod.GET)
    public ModelAndView showExcel13(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        sheet.openCellByDefinedName("testA1").setValue("Tom");
        sheet.openCellByDefinedName("testB1").setValue("John");

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        poCtrl.setSaveDataPage("save/exl/data3");
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"test13.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel13");
        return mv;
    }

    @RequestMapping(value="/excel14", method= RequestMethod.GET)
    public ModelAndView showExcel14(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        poCtrl.setCaption("简单的给Excel赋值");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTableByDefinedName("report", 10, 5, false);

        table.getDataFields().get(0).setValue("轮胎");
        table.getDataFields().get(1).setValue("100");
        table.getDataFields().get(2).setValue("120");
        table.getDataFields().get(3).setValue("500");
        table.getDataFields().get(4).setValue("120%");

        table.nextRow();

        table.close();

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        poCtrl.setSaveDataPage("/save/exl/data4");
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir+"test14.xls", OpenModeType.xlsNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel14");
        return mv;
    }



    private  PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }









}
