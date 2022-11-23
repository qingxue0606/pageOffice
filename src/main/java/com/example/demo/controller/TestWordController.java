package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpRequest;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.ServletInputStream;
import javax.servlet.ServletRequest;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.awt.*;
import java.io.*;
import java.net.URLEncoder;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Map;

@Controller
public class TestWordController {
    @Value("${testPath}")
    private String dir;


    @RequestMapping(value = "/xiang1/ajaxtest", method = RequestMethod.POST)
    @ResponseBody
    public String ajcs(HttpSession session){

        Object user=session.getAttribute("user");
        System.out.println("urser:"+ user);



        return "ok";

    }


    @RequestMapping(value = "/xiang/pdf", method = RequestMethod.GET)
    public ModelAndView showPdf(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        //--- PageOffice的调用代码 开始 -----
        PDFCtrl poCtrl1 = new PDFCtrl(request);
        poCtrl1.setServerPage("/poserver.zz"); //此行必须
        // Create custom toolbar
        poCtrl1.addCustomToolButton("保存", "Save()", 1);
        poCtrl1.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        poCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        poCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        poCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        poCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        poCtrl1.addCustomToolButton("test", "test()", 11);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("向左旋转90度", "RotateRight()", 12);
        poCtrl1.addCustomToolButton("向右旋转90度", "RotateRight()", 13);
        //poCtrl1.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");

        //poCtrl1.setMenubar(false);
        poCtrl1.setSaveFilePage("/test/save/pdf1");
        poCtrl1.webOpen(dir +"xiang\\"+ "1103.pdf");


        map.put("pageoffice", poCtrl1.getHtmlCode("PDFCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/pdf1");
        return mv;

    }

    @RequestMapping(value = "/xiang/ppt", method = RequestMethod.GET)
    public ModelAndView showPpt(HttpSession session, HttpServletRequest request, Map<String, Object> map) throws Exception {

        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        System.out.println(request.getClass().getName());
//设置服务器页面
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");
//添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 21);
        poCtrl.addCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
        //poCtrl.setAllowCopy(false);//禁止拷贝
        //poCtrl.setMenubar(false);//隐藏菜单栏
        //poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        //poCtrl.setCustomToolbar(false);//隐藏自定义工具栏

        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //poCtrl.setAllowCopy(false);


//设置保存页面
        poCtrl.setSaveFilePage("/test/save/ppt1");

        poCtrl.webOpen(dir +"xiang\\"+ "0210.ppt",OpenModeType.pptNormalEdit,"张佚名");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/ppt1");
        return mv;

    }



    @RequestMapping(value = "/xiang/word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpSession session, HttpServletRequest request, Map<String, Object> map) {
        Object id=request.getParameter("id");

        System.out.println("id:"+id);
        Object name=request.getParameter("name");
        System.out.println("name:"+name);

        Object user=session.getAttribute("user");
        //System.out.println("user:"+user);


        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        System.out.println(request.getClass().getName());
//设置服务器页面
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");


        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/test/save/doc1");//设置保存的action
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //poCtrl.setTimeSlice(10);
        //poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
        //poCtrl.setSaveDataPage("/test/save/doc1-2");
        //poCtrl.setSaveFilePage("/test/save/doc1");
        //poCtrl.setCustomToolbar(false);


        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义1", "Test1", 3); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("test", "Test", 5); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义2", "Test2", 4); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义3", "Test3", 6); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义4", "Test4", 7); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义5", "Test5", 8); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 21);
        //poCtrl.setAllowCopy(false);//禁止拷贝



        WordDocument doc = new WordDocument();
        //doc.setEnableAllDataRegionsEditing(true);
        doc.openDataRegion("PO_chapter11").setValue("[word]" + dir + "xiang\\" +"1013-2.docx"  + "[/word]");


        //dataRegion.setValue("[image width=70 height=22]/images/logo.jpg[/image]");
        //dataRegion.setValue("[image ]/images/logo.jpg[/image]");






        //dataRegion.setSubmitAsFile(true);
       //dataRegion.setValue("[word]" + dir + "xiang\\" +"0613.doc"  + "[/word]");

        /*Table table1 = dataRegion.
                createTable(5, 6, WdAutoFitBehavior.wdAutoFitWindow);
        int i = 1;
        //table1.setPreferredWidth(300.0f);
        while (i <= 5) {
            table1.openCellRC(i, 2).setValue("A" + String.valueOf(i));
            table1.openCellRC(i, 3).setValue("B" + String.valueOf(i));
            table1.openCellRC(i, 4).setValue("C" + String.valueOf(i));
            table1.openCellRC(i, 5).setValue("D" + String.valueOf(i));
            i++;
        }
        table1.openCellRC(1,1).getShading().setBackgroundPatternColor(Color.red);*/
        //dataRegion.setEditing(true);
        /*table1.openColumn(1).setWidth(50.2f);
        table1.openColumn(2).setWidth(50.2f,WdRulerStyle.wdAdjustSameWidth);
        table1.openColumn(3).setWidth(50.2f);
        table1.openColumn(4).setWidth(50.2f);
        table1.openColumn(5).setWidth(50.2f);
        table1.openColumn(6).setWidth(50.2f);*/
        //table1.setPreferredWidthType(WdPreferredWidthType.wdPreferredWidthPoints);


/*
        DataRegion dataRegion=doc.openDataRegion("PO_Content");
        dataRegion.setValue("2022");
        DataRegion dataRegion2=doc.openDataRegion("PO_xiang2");
        dataRegion2.openTable(1).openCellRC(1,1).getBorder().setBorderType(WdBorderType.wdDiagonalUp);

*/

        //poCtrl.setWriter(doc);
        //poCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");



        //poCtrl.setAllowCopy(false);


        //poCtrl.setProtectPassword("000000");
        //poCtrl.setEnableUserProtection(true);
        //poCtrl.setZoomSealServer("http://10.61.0.42:8080/ZoomSealEnt/enserver.zz");
 

        poCtrl.webOpen(dir+"xiang\\"+"1117.docx", OpenModeType.docNormalEdit,"1381818181818");
        //poCtrl.webOpen(dir+"xiang\\"+"444.docx", OpenModeType.docNormalEdit,"张三");
        //poCtrl.webOpen("/sa00sf/dfa", OpenModeType.docNormalEdit,"张三");
        //poCtrl.webOpen("/word/xiang1.docx", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/Word");
        return mv;
    }


    /**
     * 后台保存文件
     * @param session
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/xiang/word2", method = RequestMethod.GET)
    public ModelAndView showWord2(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("xiang/Word2");
        return mv;
    }
    /**
     * 打开后运行宏
     * @param session
     * @param request
     * @param map
     * @return
     */

    @RequestMapping(value = "/xiang/word3", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpSession session, HttpServletRequest request, Map<String, Object> map) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/poserver.zz");//设置授权程序servlet


        fmCtrl.setSaveFilePage("/test/save/doc2");
        WordDocument wordDocument=new WordDocument();

        //wordDocument.getWaterMark().setText("xiang");


        //打开数据区域
        //DataRegion dataRegion = wordDocument.openDataRegion("PO_TPCS1");
        //dataRegion.setValue("[image]" + dir + "xiang\\" + "0520.jpg[/image]");


        fmCtrl.setWriter(wordDocument);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        


        //fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");

        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument(dir+"xiang\\"+"444.docm", DocumentOpenType.Word);



        map.put("pageoffice",fmCtrl.getHtmlCode("PageOfficeCtrl1"));


        ModelAndView mv = new ModelAndView("xiang/Word3");
        return mv;
    }

    @RequestMapping(value = "/xiang/word4", method = RequestMethod.GET)
    public ModelAndView showWord4(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        int id=Integer.parseInt(request.getParameter("id"));
        System.out.println(id);
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);
        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        poCtrl.setSaveFilePage("/test/save/doc3");//设置保存的action



        String filePath="1.doc";
        switch (id){
            case 1:
                filePath="1.doc";
                break;
            case 2:
                filePath="2.doc";
                break;
            case 3:
                filePath="3.doc";
                break;
        }


        poCtrl.webOpen(dir+"xiang\\test4\\"+filePath, OpenModeType.docAdmin,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/Word4");
        return mv;
    }


    @RequestMapping(value = "/xiang/word5", method = RequestMethod.GET)
    public ModelAndView showWord5(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("xiang/Word5");
        return mv;
    }



    @RequestMapping(value = "/xiang/word11", method = RequestMethod.GET)
    public ModelAndView showWord11(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("xiang/Word11");
        return mv;
    }


    @RequestMapping(value = "/xiang/word6", method = RequestMethod.GET)
    public ModelAndView showWord6(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/test/save/doc1");//设置保存的action
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮

        poCtrl.addCustomToolButton("Test2", "Test2", 4); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("Test3", "Test3", 5); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 6); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 7); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 8); //添加自定义盖章按钮
        poCtrl.setOfficeVendor(OfficeVendorType.AutoSelect);//自动选择
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");

        //poCtrl.setOfficeVendor(OfficeVendorType.WPSOffice);
        poCtrl.webOpen(dir+"xiang\\"+"test6.doc", OpenModeType.docRevisionOnly,"张三");
        //poCtrl.webOpen(dir+"xiang\\"+"计量-压降仪-0008-2019-国网舟山供电公司计量中心-PT3000-070675.docx", OpenModeType.docNormalEdit,"张三");
        //poCtrl.webOpen("/1.doc", OpenModeType.docAdmin, "张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/Word6");
        return mv;
    }

    @RequestMapping(value = "/xiang/word9", method = RequestMethod.GET)
    public ModelAndView showWord9(HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("/xiang/Word9");
        return mv;
    }



    @RequestMapping(value = "/xiang/word10", method = RequestMethod.GET)
    public ModelAndView showWord10(HttpServletRequest request, Map<String, Object> map) {
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        //fmCtrl.setSaveFilePage("/save/doc/data20");
        fmCtrl.setSaveFilePage("/test/save/doc4");//设置保存的action.setSaveFilePage("/test/save/doc1");//设置保存的action
        //fmCtrl.setWriter(doc);

        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.fillDocumentAsPDF(dir+"xiang\\"+"0416.doc", DocumentOpenType.Word, "a.pdf");


        map.put("pageoffice", fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/xiang/Word10");
        return mv;
    }





    @RequestMapping("/test/save/ppt1")
    public void savePpt1(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);

        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.close();

    }




    @RequestMapping("/test/save/doc1")
    public void   saveDoc1(HttpServletRequest request, HttpServletResponse response) throws IOException {



        FileSaver fs = new FileSaver(request, response);
        fs.getFileStream();
        request.getRequestURL();
        //System.out.println(fs.getDocumentText());
        System.out.println("fs.getFileSize()-->"+fs.getFileSize());


        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        //fs.getFileExtName();





        //System.out.println("222:"+fs.getDocumentText().length());
        //fs.setCustomSaveResult(URLEncoder.encode( "项", "UTF-8" ));

        fs.close();


    }
    @RequestMapping("/test/save/doc1-2")
    public void saveDoc1_2(HttpServletRequest request, HttpServletResponse response) throws IOException {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        byte[] bytes = null;

        bytes = doc.openDataRegion("PO_fund_performance_mode").getFileBytes();

        doc.close();
        //Resource resource = new ClassPathResource("static/word/" + filePath+"D:\\project\\pageOffice\\test\\xiang");
        Resource resource = new ClassPathResource("D:\\xiang.doc");
        File file = new File("D:\\xiang.doc");
        //filePath = request.getSession().getServletContext().getRealPath("SetDrByUserWord2/doc/") + "/" + filePath;
        FileOutputStream outputStream = new FileOutputStream(file);
        outputStream.write(bytes);
        outputStream.flush();
        outputStream.close();


    }


    /**
     * 保存后台文件
     */

    @RequestMapping("/test/save/doc2")
    public void saveDoc2(HttpServletRequest request, HttpServletResponse response) throws UnsupportedEncodingException {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "xiang\\" + "test5.doc");

        fs.setCustomSaveResult(java.net.URLEncoder.encode("ok","utf-8"));

        fs.close();

    }

    @RequestMapping("/test/save/doc3")
    public void saveDoc3(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);

        fs.saveToFile(dir + "xiang\\test4\\" + fs.getFileName());

        fs.setCustomSaveResult("ok");
        fs.close();
    }

    @RequestMapping("/test/save/doc4")
    public void saveDoc4(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);


        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.setCustomSaveResult("ok");

        fs.close();



    }
    @RequestMapping("/test/save/pdf1")
    public void savePdf1(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);


        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.setCustomSaveResult("ok");
        fs.close();



    }


    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
