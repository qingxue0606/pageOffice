package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
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
import java.awt.*;
import java.io.UnsupportedEncodingException;
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
        poCtrl1.addCustomToolButton("打印", "Print()", 6);
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
        poCtrl1.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        poCtrl1.setAllowCopy(false);

        poCtrl1.webOpen(dir +"xiang\\"+ "11土方路基现场质量检验报告单.pdf");

        map.put("pageoffice", poCtrl1.getHtmlCode("PDFCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/pdf1");
        return mv;

    }

    @RequestMapping(value = "/xiang/ppt", method = RequestMethod.GET)
    public ModelAndView showPpt(HttpSession session, HttpServletRequest request, Map<String, Object> map) {

        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
//设置服务器页面
        poCtrl.setServerPage(request.getContextPath()+"/poserver.zz");
//添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.addCustomToolButton("关闭","Close",21);
        poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 21);
        poCtrl.addCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
//设置保存页面
        poCtrl.setSaveFilePage("/test/save/ppt1");
        poCtrl.setAllowCopy(false);

        poCtrl.webOpen(dir +"xiang\\"+ "test.pptx",OpenModeType.pptReadOnly,"张佚名");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/ppt1");
        return mv;

    }



    @RequestMapping(value = "/xiang/word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpSession session, HttpServletRequest request, Map<String, Object> map) {
        Object a=request.getParameter("id");

        System.out.println(a);
        Object b=request.getParameter("name");
        System.out.println(b);

        Object user=session.getAttribute("user");
        System.out.println("user:"+user);
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/test/save/doc1");//设置保存的action
        //poCtrl.setSaveDataPage("/test/save/doc1-2");
        poCtrl.setCaption("AAAA&&#38;BBB");
        //poCtrl.setCustomToolbar(false);


        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义1", "Test1.show()", 3); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("test", "Test", 5); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义2", "Test2", 4); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义3", "Test3", 6); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义4", "Test4", 7); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义5", "Test5", 8); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 21);
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");




        //poCtrl.setTheme(ThemeType.CustomStyle);
        //poCtrl.setAllowCopy(false);

        poCtrl.setCaption("11111");
        poCtrl.setFileTitle("1");
        //poCtrl.setAllowCopy(false);

        poCtrl.setJsFunction_AfterDocumentSaved("AfterDocumentSaved()");





        //poCtrl.setCaption("项");
        //poCtrl.setTitlebar(false); //隐藏标题栏
        //poCtrl.setMenubar(false); //隐藏菜单栏
        //poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        //poCtrl.setCustomToolbar(false);//隐藏自定义工具栏





        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_xiang");

        dataRegion1.setValue("1111");
        dataRegion1.setEditing(true);

        DataRegion mydr1 = doc.createDataRegion("PO_first", DataRegionInsertType.After, "PO_xiang");
        DataRegion mydr2 = doc.createDataRegion("PO_second", DataRegionInsertType.After, "PO_first");
        mydr1.setValue("\n\r");
        mydr2.setValue("");

        mydr2.selectEnd();
        doc.insertPageBreak();//插入分页符








        DataRegion dataRegion2 = doc.openDataRegion("PO_xiang2");

        //dataRegion2.setValue("[image]/images/img_1.jpg[/image]");


        dataRegion2.setEditing(true);
        DataRegion dataRegion3 = doc.openDataRegion("PO_xiang3");

        dataRegion3.setValue("");


        dataRegion3.setEditing(true);
        //DataRegion dataRegion2 = doc.openDataRegion("PO_xiang");
        //dataRegion2.setValue("台风力气吗测试\r\n\r\n1\r\n\r\n1\n");

        //dataRegion2.setEditing(true);

        //DataTag deptTag = doc.openDataTag("{Tag}");
        //给DataTag对象赋值
        //deptTag.setValue("B部门");
        //doc.getWaterMark().setText("xiang");


        //dataRegion2.setValue("市场");
        //dataRegion2.getFont().setColor(Color.orange);





        //poCtrl.setWriter(doc);



        //poCtrl.setOfficeVendor(OfficeVendorType.WPSOffice);

        poCtrl.webOpen(dir+"xiang\\"+"zw_20200109.doc", OpenModeType.docNormalEdit,"张三");
        //poCtrl.webOpen("/sasf/dfa", OpenModeType.docSubmitForm,"张三");
        //poCtrl.webOpen("/0709.doc", OpenModeType.docSubmitForm, "张三");
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

        wordDocument.getWaterMark().setText("xiang");


        //打开数据区域
        DataRegion dataRegion = wordDocument.openDataRegion("PO_regTable");
        dataRegion.setValue("xiang");


        fmCtrl.setWriter(wordDocument);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");


        //fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");

        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument(dir+"xiang\\"+"test4.doc", DocumentOpenType.Word);



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


        poCtrl.webOpen(dir+"xiang\\test4\\"+filePath, OpenModeType.docNormalEdit,"张三");
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




    @RequestMapping("/test/save/ppt1")
    public void savePpt1(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.close();

    }




    @RequestMapping("/test/save/doc1")
    public void   saveDoc1(HttpServletRequest request, HttpServletResponse response) {

        request.getParameter("id");
        System.out.println(request.getParameter("id"));


        FileSaver fs = new FileSaver(request, response);
        System.out.println(fs.getDocumentText());
        System.out.println("fs.getFileSize()-->"+fs.getFileSize());
        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        System.out.println(fs.getFileName());
        fs.setCustomSaveResult("setCustomSaveResult");

        fs.close();

    }
    @RequestMapping("/test/save/doc1-2")
    public void saveDoc1_2(HttpServletRequest request, HttpServletResponse response) {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);

        //获取提交的数值
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion Name1 = doc.openDataRegion("PO_year1");
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion Name2 = doc.openDataRegion("PO_result11");
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion Name3 = doc.openDataRegion("PO_public1");

        System.out.println(Name1.getValue());
        System.out.println(Name2.getValue());
        System.out.println(Name3.getValue());

        doc.close();

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


    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
