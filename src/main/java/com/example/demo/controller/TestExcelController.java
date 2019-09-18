package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataTag;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Map;

@Controller
public class TestExcelController {
    @Value("${testPath}")
    private String dir;



    @RequestMapping(value = "/xiang/excel", method = RequestMethod.GET)
    public ModelAndView showExcel(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("导入文件", "importData()", 16);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("test", "test()", 1);
        poCtrl.setJsFunction_AfterDocumentOpened( "AfterDocumentOpened()");



        //poCtrl.setSaveDataPage("/test/save/ex1");
        poCtrl.setSaveFilePage("/test/save/ex1");

        //poCtrl.setSaveDataPage("/test/save/ex2");

        //poCtrl.setJsFunction_OnExcelCellClick("OnCellClick()");
        //定义Workbook对象
        Workbook workBook = new Workbook();
        //定义Sheet对象，"Sheet1"是打开的Excel表单的名称
        Sheet sheet = workBook.openSheet("Z01-收入支出决算总表");
        //定义table对象，设置table对象的设置范围
        //com.zhuozhengsoft.pageoffice.excelwriter.Table table = sheet.openTable("B4:F13");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        //table.setSubmitName("Info");
        Cell cell=sheet.openCellRC(10,5);
        cell.setValue("eeeee");
        cell.setSubmitName("eeeee2");
        cell.setReadOnly(false);


        Sheet sheet2 = workBook.openSheet("Z01-1-财政拨款收入支出决算总表");
        //定义table对象，设置table对象的设置范围
        //com.zhuozhengsoft.pageoffice.excelwriter.Table table = sheet.openTable("B4:F13");
        //设置table对象的提交名称，以便保存页面获取提交的数据
        //table.setSubmitName("Info");
        Cell cell2=sheet2.openCellRC(3,4);
        cell2.setValue("eeeee");
        cell2.setSubmitName("eeeee2");





        /*com.zhuozhengsoft.pageoffice.excelwriter.Cell cell= sheet.openCell("C1") ;
        cell.setValue("卓正软件"); // 此单元格不会提交
        cell.setReadOnly(true);*/


        //定义table对象，设置table对象的设置范围
        /*com.zhuozhengsoft.pageoffice.excelwriter.Table table = sheet.openTable("A17:J21");

        //设置table对象的提交名称，以便保存页面获取提交的数据
        table.setSubmitName("Info");



        for(int i=0; i < 10; i++)
        {
            table.getDataFields().get(0).setValue("产品 ");
            table.getDataFields().get(1).setValue("100");
            table.getDataFields().get(2).setValue("3列");
            table.nextRow();
        }
        table.close();*/


        //poCtrl.setWriter(workBook);

        poCtrl.webOpen(dir+"xiang\\"+"32.xls",OpenModeType.xlsNormalEdit,"张佚名");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/Excel");
        return mv;

    }









    @RequestMapping("/test/save/ex1")
    public void saveEx1(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.close();



    }
    @RequestMapping("/test/save/ex2")
    public void saveEx2(HttpServletRequest request, HttpServletResponse response) {

        com.zhuozhengsoft.pageoffice.excelreader.Workbook workBook = new com.zhuozhengsoft.pageoffice.excelreader.Workbook(request, response);
        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet = workBook.openSheet("Z01-收入支出决算总表");


        String c14=sheet.openCell("E10").getValue();
        System.out.println("E10:"+c14);


        com.zhuozhengsoft.pageoffice.excelreader.Sheet sheet2 = workBook.openSheet("Z01-1-财政拨款收入支出决算总表");

        String c214=sheet2.openCell("E10").getValue();
        System.out.println("E10:"+c214);



        ArrayList cells = sheet.getCells();
        for (Object o:cells) {
            com.zhuozhengsoft.pageoffice.excelreader.Cell cell=(com.zhuozhengsoft.pageoffice.excelreader.Cell)o;
            String s1= cell.getSubmitName();
            System.out.println("SubmitName "+s1);
            System.out.println("getValue "+cell.getValue());
            System.out.println("getText "+cell.getText());
        }
        System.out.println("个数"+cells.size());

        //Table table = sheet.openTable("Info");
        /*Table table = sheet.openTableBySubmitName("qqq");

        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "1列"
                        + table.getDataFields().get(0).getText();
                content += " 2列："
                        + table.getDataFields().get(1).getText();
                content += " 3列"
                        + table.getDataFields().get(2).getText();

                content += "\n";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();*/

        workBook.showPage(500, 400);
        workBook.close();
        //System.out.println(content);

    }





    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
