package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataTag;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class WordController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value = "/word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/save");//设置保存的action
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        //新建一个WordDocument用来操作数据
        WordDocument doc = new WordDocument();

        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("☑");

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");

        poCtrl.setWriter(doc);


        poCtrl.webOpen(dir + "test1.doc", OpenModeType.docAdmin, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word");
        return mv;
    }

    @RequestMapping(value = "/word2", method = RequestMethod.GET)
    public ModelAndView showWord2(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        //poCtrl.setSaveFilePage("/save/doc2");//设置保存的action
        poCtrl.setSaveDataPage("/save/doc2");
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮

        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
        //设置DataRegion的可编辑性
        dataRegion1.setEditing(true);

        //为DataRegion赋值,此处的值可在页面中打开Word文档后自己进行修改
        dataRegion1.setValue("");

        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);
        dataRegion2.setValue("");

        poCtrl.setWriter(wordDoc);


        poCtrl.webOpen(dir + "test2.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word");
        return mv;
    }


    @RequestMapping(value = "/word3", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //poCtrl.setSaveFilePage("/save/doc2");//设置保存的action
        poCtrl.setSaveDataPage("/save");
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮


        //poCtrl.setCustomToolbar(false);
        //poCtrl.setOfficeToolbars(false);
        //poCtrl.setAllowCopy(false);//禁止拷贝

        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");


        poCtrl.webOpen(dir + "sdasd.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word3");
        return mv;
    }

    @RequestMapping(value = "/word4", method = RequestMethod.GET)
    public ModelAndView showWord4(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion = doc.openDataRegion("PO_regTable");
        //打开table，openTable(index)方法中的index代表Word文档中table位置的索引，从1开始
        Table table = dataRegion.openTable(1);

        //给table中的单元格赋值， openCellRC(int,int)中的参数分别代表第几行、第几列，从1开始
        table.openCellRC(3, 1).setValue("A公司");
        table.openCellRC(3, 2).setValue("开发部");
        table.openCellRC(3, 3).setValue("李清");

        //插入一行，insertRowAfter方法中的参数代表在哪个单元格下面插入一个空行
        table.insertRowAfter(table.openCellRC(3, 3));

        table.openCellRC(4, 1).setValue("B公司");
        table.openCellRC(4, 2).setValue("销售部");
        table.openCellRC(4, 3).setValue("张三");

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir + "test4.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word3");
        return mv;
    }


    @RequestMapping(value = "/word5", method = RequestMethod.GET)
    public ModelAndView showWord5(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //定义WordDocument对象
        WordDocument doc = new WordDocument();

        //定义DataTag对象
        DataTag deptTag = doc.openDataTag("{部门名}");
        deptTag.setValue("技术");

        DataTag userTag = doc.openDataTag("{姓名}");
        userTag.setValue("李四");

        DataTag dateTag = doc.openDataTag("【时间】");
        dateTag.setValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()).toString());

        poCtrl.setWriter(doc);
        //打开Word文件

        poCtrl.addCustomToolButton("测试按钮", "myTest", 0);

        // 设置文件打开后执行的js function
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");


        poCtrl.webOpen(dir + "test5.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word4");
        return mv;
    }


    @RequestMapping(value = "/word6", method = RequestMethod.GET)
    public ModelAndView showWord6(HttpServletRequest request, Map<String, Object> map) {

        String userName = "somebody";

        String userId = request.getParameter("userid").toString();
        if (userId.equals("1")) {
            userName = "张三";
        } else {
            userName = "李四";
        }

        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setSaveFilePage("/save");

        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        dataRegion1.setEditing(true);



        poCtrl.setWriter(doc);

        //设置并发控制时间
        poCtrl.setTimeSlice(20);


        poCtrl.webOpen(dir + "test6.doc", OpenModeType.docSubmitForm, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word6");
        return mv;
    }


    @RequestMapping(value = "/word7", method = RequestMethod.GET)
    public ModelAndView showWord7(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //添加自定义按钮
        poCtrl.addCustomToolButton("另存HTML", "saveAsHTML", 1);
        poCtrl.setSaveFilePage("/save/doc7");
        //设置并发控制时间
        poCtrl.setTimeSlice(20);


        poCtrl.webOpen(dir + "test7.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word7");
        return mv;
    }

    @RequestMapping(value = "/word8", method = RequestMethod.GET)
    public ModelAndView showWord8(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //添加自定义按钮
        poCtrl.addCustomToolButton("另存MHT", "saveAsMHT", 1);
        poCtrl.setSaveFilePage("/save/doc8");
        //设置并发控制时间
        poCtrl.setTimeSlice(20);


        poCtrl.webOpen(dir + "test8.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word8");
        return mv;
    }

    @RequestMapping(value = "/word9", method = RequestMethod.GET)
    public ModelAndView showWord9(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        // 设置文件保存之前执行的事件
        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");
// 设置文件保存之后执行的事件
        poCtrl.setJsFunction_AfterDocumentSaved("AfterDocumentSaved()");
        poCtrl.setSaveFilePage("/save/doc7");
        //设置并发控制时间
        poCtrl.setTimeSlice(20);


        poCtrl.webOpen(dir + "test9.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word9");
        return mv;
    }

    @RequestMapping(value = "/word10", method = RequestMethod.GET)
    public ModelAndView showWord10(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("关闭", "Close", 21);

        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");


        poCtrl.setSaveFilePage("/save/doc7");


        poCtrl.webOpen(dir + "test10.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word10");
        return mv;
    }

    @RequestMapping(value = "/word11", method = RequestMethod.GET)
    public ModelAndView showWord11(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        WordDocument wordDoc = new WordDocument();

        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_userName");
        //设置DataRegion的可编辑性
        dataRegion1.setEditing(true);
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_deptName");
        dataRegion2.setEditing(true);
        poCtrl.setWriter(wordDoc);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //设置保存数据的页面


        poCtrl.setSaveFilePage("/save/doc7");
        poCtrl.setSaveDataPage("/save/doc/data");


        poCtrl.webOpen(dir + "test11.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word11");
        return mv;
    }

    @RequestMapping(value = "/word12", method = RequestMethod.GET)
    public ModelAndView showWord12(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        poCtrl.addCustomToolButton("导入文件", "importData()", 15);
        poCtrl.addCustomToolButton("提交数据", "submitData()", 1);
        WordDocument doc = new WordDocument();
        poCtrl.setWriter(doc);


        poCtrl.setSaveDataPage("/save/doc/data12");


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word12");
        return mv;
    }


    @RequestMapping(value = "/word13", method = RequestMethod.GET)
    public ModelAndView showWord13(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        WordDocument wordDoc = new WordDocument();
        wordDoc.setDisableWindowRightClick(true);//禁止word鼠标右键
        //wordDoc.setDisableWindowDoubleClick(true);//禁止word鼠标双击
        //wordDoc.setDisableWindowSelection(true);//禁止在word中选择文件内容

        poCtrl.setWriter(wordDoc);
        //打开文件
        poCtrl.webOpen(dir + "test11.doc", OpenModeType.docSubmitForm, "张三");


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word12");
        return mv;
    }

    @RequestMapping(value = "/word15", method = RequestMethod.GET)
    public ModelAndView showWord15(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);

        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        poCtrl.setSaveFilePage("/save");//设置保存的action
        //打开文件
        poCtrl.webOpen(dir + "test15.doc", OpenModeType.docSubmitForm, "张三");


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word15");
        return mv;
    }


    @RequestMapping(value = "/word16", method = RequestMethod.GET)
    public ModelAndView showWord16(HttpServletRequest request, Map<String, Object> map) {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("打印设置", "PrintSet", 0);
        poCtrl.addCustomToolButton("打印", "PrintFile", 6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("关闭", "Close", 21);
        poCtrl.setSaveFilePage("/save");//设置保存的action
        //打开文件
        poCtrl.webOpen(dir + "test16.doc", OpenModeType.docSubmitForm, "张三");


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word16");
        return mv;
    }

    @RequestMapping(value = "/word17", method = RequestMethod.GET)
    public ModelAndView showWord17(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("隐藏痕迹", "hideRevision", 18);
        poCtrl.addCustomToolButton("显示痕迹", "showRevision", 9);
        //设置保存页面
        poCtrl.setSaveFilePage("/save");//设置保存的action
        //打开文件
        poCtrl.webOpen(dir + "test17.doc", OpenModeType.docSubmitForm, "张三");


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word17");
        return mv;
    }


    @RequestMapping(value = "/word18", method = RequestMethod.GET)
    public ModelAndView showWord18(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        poCtrl.setAllowCopy(false);//禁止拷贝
        poCtrl.setMenubar(false);//隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //设置页面的显示标题

        poCtrl.setCaption("演示：文件在线安全浏览");
        //打开文件
        poCtrl.webOpen(dir + "test.ppt", OpenModeType.pptReadOnly, "张三");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word18");
        return mv;
    }


    @RequestMapping(value = "/word20", method = RequestMethod.GET)
    public ModelAndView showWord20(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.setSaveFilePage("/save/doc/data13?id=1");


        poCtrl.webOpen("/openWord?id=1", OpenModeType.docNormalEdit, "张三");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word20");
        return mv;
    }

    @RequestMapping(value = "/word22", method = RequestMethod.GET)
    public ModelAndView showWord21(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("另存为PDF文件", "SaveAsPDF()", 1);
//设置保存页面
        poCtrl.setSaveFilePage("/save/doc/data15");
        String fileName = "test22.doc";
        String pdfName = fileName.substring(0, fileName.length() - 4) + ".pdf";


        poCtrl.webOpen(dir + "其他类型-多.docx", OpenModeType.docNormalEdit, "张三");
        //poCtrl.webOpen(dir + "test2.xls", OpenModeType.xlsNormalEdit, "张三");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("pdfName", pdfName);

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word22");
        return mv;
    }

    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
