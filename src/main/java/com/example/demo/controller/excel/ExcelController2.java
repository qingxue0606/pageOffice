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


    @RequestMapping(value = "/excel15", method = RequestMethod.GET)
    public ModelAndView showExcel15(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        String tempFileName = request.getParameter("temp");
        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象，参数“report”就是Excel模板中定义的单元格区域的名称
        Table table = sheet.openTableByDefinedName("report11", 10, 5, true);
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
        Calendar c = new GregorianCalendar();
        int year = c.get(Calendar.YEAR);//获取年份
        cellYear.setValue(year + "年");

        Cell cellName = sheet.openCellByDefinedName("name");
        cellName.setValue("张三");

        poCtrl.setWriter(workBook);

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir + "exl15//" + tempFileName, OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel15");
        return mv;
    }

    @RequestMapping(value = "/excel16", method = RequestMethod.GET)
    public ModelAndView showExcel16(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        //poCtrl1.setSaveDataPage("SaveData.jsp");
        //poCtrl1.addCustomToolButton("保存", "Save()", 1);
        //打开Word文件


        poCtrl.webOpen(dir + "test16.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value = "/excel17", method = RequestMethod.GET)
    public ModelAndView showExcel17(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象
        Table table = sheet.openTable("B4:F11");

        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount; i++) {
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue(i + "月");
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
        for (int i = 1; i <= rowCount2; i++) {
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue(i + "季度");
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


        poCtrl.webOpen(dir + "test16.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value = "/excel18", method = RequestMethod.GET)
    public ModelAndView showExcel18(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.setCaption("简单的给Excel赋值");

        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Sheet1");

        //定义Table对象，参数“report1”为Excel中定义的名称，“4”为名称指定区域的行数，
        //“5”为名称指定区域的列数，“true”表示表格会按实际数据行数自动扩展
        Table table = sheet.openTableByDefinedName("report", 4, 5, true);

        int rowCount = 12;//假设将要自动填充数据的实际记录条数为12
        for (int i = 1; i <= rowCount; i++) {
            //给区域中的单元格赋值
            table.getDataFields().get(0).setValue(i + "月");
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
        for (int i = 1; i <= rowCount2; i++) {
            //给区域中的单元格赋值
            table2.getDataFields().get(0).setValue(i + "季度");
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


        poCtrl.webOpen(dir + "test16.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value = "/excel19", method = RequestMethod.GET)
    public ModelAndView showExcel19(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.getRibbonBar().setTabVisible("TabHome", true);//开始
        poCtrl.getRibbonBar().setTabVisible("TabFormulas", false);//公式
        poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
        poCtrl.getRibbonBar().setTabVisible("TabData", false);//数据
        poCtrl.getRibbonBar().setTabVisible("TabReview", false);//审阅
        poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图
        poCtrl.getRibbonBar().setTabVisible("TabPageLayoutExcel", false);//页面布局

        poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮

        poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板


        poCtrl.webOpen(dir + "test16.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value = "/excel20", method = RequestMethod.GET)
    public ModelAndView showExcel20(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.setCustomToolbar(false);
        Workbook wb = new Workbook();
        Sheet sheet1 = wb.openSheet("Sheet1");
        //设置当工作表只读时，是否允许用户手动调整行列。
        sheet1.setAllowAdjustRC(true);
        poCtrl.setWriter(wb);//此行必须


        poCtrl.webOpen(dir + "test16.xls", OpenModeType.xlsReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel20");
        return mv;
    }

    @RequestMapping(value = "/excel21", method = RequestMethod.GET)
    public ModelAndView showExcel21(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        Workbook workBook = new Workbook();
        Sheet sheet1 = workBook.openSheet("Sheet1");
        //sheet1.openCell("A1").setValue("[image]images/logo.jpg[/image]");
        sheet1.openCell("A1").setValue("[image]" + dir + "xiang" + "\\111.jpg[/image]");
        poCtrl.setWriter(workBook);//此行必须


        poCtrl.webOpen(dir + "test21.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel16");
        return mv;
    }

    @RequestMapping(value = "/excel22", method = RequestMethod.GET)
    public ModelAndView showExcel22(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);


        poCtrl.webOpen(dir + "test22.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel22");
        return mv;
    }

    @RequestMapping(value = "/excel23", method = RequestMethod.GET)
    public ModelAndView showExcel23(HttpServletRequest request, Map<String, Object> map) {
        Workbook wb = new Workbook();
        // 设置背景
        Table backGroundTable = wb.openSheet("Sheet1").openTable("A1:P200");
        backGroundTable.getBorder().setLineColor(Color.white);

        // 设置标题
        wb.openSheet("Sheet1").openTable("A1:H2").merge();
        wb.openSheet("Sheet1").openTable("A1:H2").setRowHeight(30);
        Cell A1 = wb.openSheet("Sheet1").openCell("A1");
        A1.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        A1.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        A1.setForeColor(new Color(0, 128, 128));
        A1.setValue("出差开支预算");

        //设置字体
        wb.openSheet("Sheet1").openTable("A1:A1").getFont().setBold(true);
        wb.openSheet("Sheet1").openTable("A1:A1").getFont().setSize(25);

        // 画表头
        Border C4Border = wb.openSheet("Sheet1").openTable("C4:C4").getBorder();
        C4Border.setWeight(XlBorderWeight.xlThick);
        C4Border.setLineColor(Color.yellow);

        Table titleTable = wb.openSheet("Sheet1").openTable("B4:H5");
        titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);
        titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
        titleTable.getBorder().setLineColor(new Color(0, 128, 128));

        // 画表体
        Table bodyTable = wb.openSheet("Sheet1").openTable("B6:H15");
        bodyTable.getBorder().setLineColor(Color.gray);
        bodyTable.getBorder().setWeight(XlBorderWeight.xlHairline);

        Border B7Border = wb.openSheet("Sheet1").openTable("B7:B7").getBorder();
        B7Border.setLineColor(Color.white);

        Border B9Border = wb.openSheet("Sheet1").openTable("B9:B9").getBorder();
        B9Border.setBorderType(XlBorderType.xlBottomEdge);
        B9Border.setLineColor(Color.white);

        Border C6C15BorderLeft = wb.openSheet("Sheet1").openTable("C6:C15").getBorder();
        C6C15BorderLeft.setLineColor(Color.white);
        C6C15BorderLeft.setBorderType(XlBorderType.xlLeftEdge);

        Border C6C15BorderRight = wb.openSheet("Sheet1").openTable("C6:C15").getBorder();
        C6C15BorderRight.setLineColor(Color.yellow);
        C6C15BorderRight.setLineStyle(XlBorderLineStyle.xlDot);
        C6C15BorderRight.setBorderType(XlBorderType.xlRightEdge);

        Border E6E15Border = wb.openSheet("Sheet1").openTable("E6:E15").getBorder();
        E6E15Border.setLineStyle(XlBorderLineStyle.xlDot);
        E6E15Border.setBorderType(XlBorderType.xlAllEdges);
        E6E15Border.setLineColor(Color.yellow);

        Border G6G15BorderRight = wb.openSheet("Sheet1").openTable("G6:G15").getBorder();
        G6G15BorderRight.setBorderType(XlBorderType.xlRightEdge);
        G6G15BorderRight.setLineColor(Color.white);

        Border G6G15BorderLeft = wb.openSheet("Sheet1").openTable("G6:G15").getBorder();
        G6G15BorderLeft.setLineStyle(XlBorderLineStyle.xlDot);
        G6G15BorderLeft.setBorderType(XlBorderType.xlLeftEdge);
        G6G15BorderLeft.setLineColor(Color.yellow);

        Table bodyTable2 = wb.openSheet("Sheet1").openTable("B6:H15");
        bodyTable2.getBorder().setWeight(XlBorderWeight.xlThick);
        bodyTable2.getBorder().setLineColor(new Color(0, 128, 128));
        bodyTable2.getBorder().setBorderType(XlBorderType.xlAllEdges);

        // 画表尾
        Border H16H17Border = wb.openSheet("Sheet1").openTable("H16:H17").getBorder();
        H16H17Border.setLineColor(new Color(204, 255, 204));

        Border E16G17Border = wb.openSheet("Sheet1").openTable("E16:G17").getBorder();
        E16G17Border.setLineColor(new Color(0, 128, 128));

        Table footTable = wb.openSheet("Sheet1").openTable("B16:H17");
        footTable.getBorder().setWeight(XlBorderWeight.xlThick);
        footTable.getBorder().setLineColor(new Color(0, 128, 128));
        footTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

        // 设置行高列宽
        wb.openSheet("Sheet1").openTable("A1:A1").setColumnWidth(1);
        wb.openSheet("Sheet1").openTable("B1:B1").setColumnWidth(20);
        wb.openSheet("Sheet1").openTable("C1:C1").setColumnWidth(15);
        wb.openSheet("Sheet1").openTable("D1:D1").setColumnWidth(10);
        wb.openSheet("Sheet1").openTable("E1:E1").setColumnWidth(8);
        wb.openSheet("Sheet1").openTable("F1:F1").setColumnWidth(3);
        wb.openSheet("Sheet1").openTable("G1:G1").setColumnWidth(12);
        wb.openSheet("Sheet1").openTable("H1:H1").setColumnWidth(20);

        wb.openSheet("Sheet1").openTable("A16:A16").setRowHeight(20);
        wb.openSheet("Sheet1").openTable("A17:A17").setRowHeight(20);

        // 设置表格中字体大小为10
        for (int i = 0; i < 12; i++) {//excel表格行号
            for (int j = 0; j < 7; j++) {//excel表格列号
                wb.openSheet("Sheet1").openCellRC(4 + i, 2 + j).getFont().setSize(10);
            }
        }

        // 填充单元格背景颜色
        for (int i = 0; i < 10; i++) {
            wb.openSheet("Sheet1").openCell("H" + (6 + i)).setBackColor(new Color(255, 255, 153));
        }

        wb.openSheet("Sheet1").openCell("E16").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("F16").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("G16").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("E17").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("F17").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("G17").setBackColor(new Color(0, 128, 128));
        wb.openSheet("Sheet1").openCell("H16").setBackColor(new Color(204, 255, 204));
        wb.openSheet("Sheet1").openCell("H17").setBackColor(new Color(204, 255, 204));

        //填充单元格文本和公式
        Cell B4 = wb.openSheet("Sheet1").openCell("B4");
        B4.getFont().setBold(true);
        B4.setValue("出差开支预算");
        Cell H5 = wb.openSheet("Sheet1").openCell("H5");
        H5.getFont().setBold(true);
        H5.setValue("总计");
        H5.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
        Cell B6 = wb.openSheet("Sheet1").openCell("B6");
        B6.getFont().setBold(true);
        B6.setValue("飞机票价");
        Cell B9 = wb.openSheet("Sheet1").openCell("B9");
        B9.getFont().setBold(true);
        B9.setValue("酒店");
        Cell B11 = wb.openSheet("Sheet1").openCell("B11");
        B11.getFont().setBold(true);
        B11.setValue("餐饮");
        Cell B12 = wb.openSheet("Sheet1").openCell("B12");
        B12.getFont().setBold(true);
        B12.setValue("交通费用");
        Cell B13 = wb.openSheet("Sheet1").openCell("B13");
        B13.getFont().setBold(true);
        B13.setValue("休闲娱乐");
        Cell B14 = wb.openSheet("Sheet1").openCell("B14");
        B14.getFont().setBold(true);
        B14.setValue("礼品");
        Cell B15 = wb.openSheet("Sheet1").openCell("B15");
        B15.getFont().setBold(true);
        B15.getFont().setSize(10);
        B15.setValue("其他费用");

        wb.openSheet("Sheet1").openCell("C6").setValue("机票单价（往）");
        wb.openSheet("Sheet1").openCell("C7").setValue("机票单价（返）");
        wb.openSheet("Sheet1").openCell("C8").setValue("其他");
        wb.openSheet("Sheet1").openCell("C9").setValue("每晚费用");
        wb.openSheet("Sheet1").openCell("C10").setValue("其他");
        wb.openSheet("Sheet1").openCell("C11").setValue("每天费用");
        wb.openSheet("Sheet1").openCell("C12").setValue("每天费用");
        wb.openSheet("Sheet1").openCell("C13").setValue("总计");
        wb.openSheet("Sheet1").openCell("C14").setValue("总计");
        wb.openSheet("Sheet1").openCell("C15").setValue("总计");

        wb.openSheet("Sheet1").openCell("G6").setValue("  张");
        wb.openSheet("Sheet1").openCell("G7").setValue("  张");
        wb.openSheet("Sheet1").openCell("G9").setValue("  晚");
        wb.openSheet("Sheet1").openCell("G10").setValue("  晚");
        wb.openSheet("Sheet1").openCell("G11").setValue("  天");
        wb.openSheet("Sheet1").openCell("G12").setValue("  天");

        wb.openSheet("Sheet1").openCell("H6").setFormula("=D6*F6");
        wb.openSheet("Sheet1").openCell("H7").setFormula("=D7*F7");
        wb.openSheet("Sheet1").openCell("H8").setFormula("=D8*F8");
        wb.openSheet("Sheet1").openCell("H9").setFormula("=D9*F9");
        wb.openSheet("Sheet1").openCell("H10").setFormula("=D10*F10");
        wb.openSheet("Sheet1").openCell("H11").setFormula("=D11*F11");
        wb.openSheet("Sheet1").openCell("H12").setFormula("=D12*F12");
        wb.openSheet("Sheet1").openCell("H13").setFormula("=D13*F13");
        wb.openSheet("Sheet1").openCell("H14").setFormula("=D14*F14");
        wb.openSheet("Sheet1").openCell("H15").setFormula("=D15*F15");

        for (int i = 0; i < 10; i++) {
            //设置数据以货币形式显示
            wb.openSheet("Sheet1").openCell("D" + (6 + i)).setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
            wb.openSheet("Sheet1").openCell("H" + (6 + i)).setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        }

        Cell E16 = wb.openSheet("Sheet1").openCell("E16");
        E16.getFont().setBold(true);
        E16.getFont().setSize(11);
        E16.setForeColor(Color.white);
        E16.setValue("出差开支总费用");
        E16.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        Cell E17 = wb.openSheet("Sheet1").openCell("E17");
        E17.getFont().setBold(true);
        E17.getFont().setSize(11);
        E17.setForeColor(Color.white);
        E17.setFormula("=IF(C4>H16,\"低于预算\",\"超出预算\")");
        E17.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        Cell H16 = wb.openSheet("Sheet1").openCell("H16");
        H16.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        H16.setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        H16.getFont().setName("Arial");
        H16.getFont().setSize(11);
        H16.getFont().setBold(true);
        H16.setFormula("=SUM(H6:H15)");
        Cell H17 = wb.openSheet("Sheet1").openCell("H17");
        H17.setVerticalAlignment(XlVAlign.xlVAlignCenter);
        H17.setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        H17.getFont().setName("Arial");
        H17.getFont().setSize(11);
        H17.getFont().setBold(true);
        H17.setFormula("=(C4-H16)");

        // 填充数据
        Cell C4 = wb.openSheet("Sheet1").openCell("C4");
        C4.setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        C4.setValue("2500");
        Cell D6 = wb.openSheet("Sheet1").openCell("D6");
        D6.setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        D6.setValue("1200");
        wb.openSheet("Sheet1").openCell("F6").getFont().setSize(10);
        wb.openSheet("Sheet1").openCell("F6").setValue("1");
        Cell D7 = wb.openSheet("Sheet1").openCell("D7");
        D7.setNumberFormatLocal("￥#,##0.00;￥-#,##0.00");
        D7.setValue("875");
        wb.openSheet("Sheet1").openCell("F7").setValue("1");


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);
        poCtrl.setWriter(wb);

        //创建自定义菜单栏
        poCtrl.addCustomToolButton("全屏切换", "SetFullScreen()", 4);

        poCtrl.setMenubar(false);//隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具栏


        poCtrl.webOpen(dir + "test23.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("excel/excel23");
        return mv;
    }


    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
