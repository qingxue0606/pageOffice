package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class WordController3 {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value="/word38", method= RequestMethod.GET)
    public ModelAndView showWord38(HttpServletRequest request, Map<String,Object> map){
        ModelAndView mv = new ModelAndView("/word/Word38");
        return mv;
    }

    @RequestMapping(value="/word39", method= RequestMethod.GET)
    public ModelAndView showWord39(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        String userName=request.getParameter("userName");

        if ( userName.equals("zhangsan") ) userName = "张三";
        if (userName.equals("lisi")) userName = "李四";
        if (userName.equals("wangwu")) userName = "王五";

        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.setJsFunction_AfterDocumentOpened("ShowByUserName");

        //设置保存页
        poCtrl.setSaveDataPage("/save/common");



        poCtrl.webOpen(dir+"test39.doc", OpenModeType.docNormalEdit,userName);
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word39");
        return mv;
    }


    @RequestMapping(value="/word40", method= RequestMethod.GET)
    public ModelAndView showWord40(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        DataRegion dataReg = doc.openDataRegion("PO_table");
        Table table = dataReg.openTable(1);
        //合并table中的单元格
        table.openCellRC(1, 1).mergeTo(1, 4);
        //给合并后的单元格赋值
        table.openCellRC(1, 1).setValue("销售情况表");
        //设置单元格文本样式
        table.openCellRC(1, 1).getFont().setColor(Color.red);
        table.openCellRC(1, 1).getFont().setSize(24);
        table.openCellRC(1, 1).getFont().setName("楷体");
        table.openCellRC(1, 1).getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphCenter);


        poCtrl.setWriter(doc);
        poCtrl.setCustomToolbar(false);



        poCtrl.webOpen(dir+"test40.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word40");
        return mv;
    }
    @RequestMapping(value="/word41", method= RequestMethod.GET)
    public ModelAndView showWord41(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);


        WordDocument doc = new WordDocument();
        DataRegion dataReg = doc.openDataRegion("PO_deptName");
        dataReg.getShading().setBackgroundPatternColor(Color.pink);
        //dataReg.setEditing(true);
        poCtrl.setWriter(doc);
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");
        poCtrl.setOfficeToolbars(false);
        poCtrl.setCaption("为方便用户知道哪些地方可以编辑，所以设置了数据区域的背景色");
        poCtrl.setSaveFilePage("SaveFile.jsp");

        poCtrl.webOpen(dir+"test41.doc", OpenModeType.docSubmitForm,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word41");
        return mv;
    }
    @RequestMapping(value="/word42", method= RequestMethod.GET)
    public ModelAndView showWord42(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("张三");
        //设置字体样式
        dataRegion1.getFont().setColor(Color.blue);
        dataRegion1.getFont().setSize(24);
        dataRegion1.getFont().setName("隶书");
        dataRegion1.getFont().setBold(true);

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");
        dataRegion2.getFont().setColor(Color.red);

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);

        poCtrl.webOpen(dir+"test42.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word42");
        return mv;
    }

    @RequestMapping(value="/word43", method= RequestMethod.GET)
    public ModelAndView showWord43(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        doc.getTemplate().defineDataRegion("Name", "[ 姓名 ]");
        doc.getTemplate().defineDataRegion("Address", "[ 地址 ]");
        doc.getTemplate().defineDataRegion("Tel", "[ 电话 ]");
        doc.getTemplate().defineDataRegion("Phone", "[ 手机 ]");
        doc.getTemplate().defineDataRegion("Sex", "[ 性别 ]");
        doc.getTemplate().defineDataRegion("Age", "[ 年龄 ]");
        doc.getTemplate().defineDataRegion("Email", "[ 邮箱 ]");
        doc.getTemplate().defineDataRegion("QQNo", "[ QQ号 ]");
        doc.getTemplate().defineDataRegion("MSNNo", "[ MSN号 ]");

        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("定义数据区域", "ShowDefineDataRegions()", 3);

        poCtrl.setSaveFilePage("/save/common");

        poCtrl.setTheme(ThemeType.Office2007);
        poCtrl.setBorderStyle(BorderStyleType.BorderThin);
        poCtrl.setWriter(doc);

        poCtrl.webOpen(dir+"test43.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word43");
        return mv;
    }
    @RequestMapping(value="/word44", method= RequestMethod.GET)
    public ModelAndView showWord44(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        doc.getTemplate().defineDataTag("{ 甲方 }");
        doc.getTemplate().defineDataTag("{ 乙方 }");
        doc.getTemplate().defineDataTag("{ 担保人 }");
        doc.getTemplate().defineDataTag("【 合同日期 】");
        doc.getTemplate().defineDataTag("【 合同编号 】");



        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("定义数据标签", "ShowDefineDataTags()", 20);

        poCtrl.setTheme(ThemeType.Office2007);
        poCtrl.setBorderStyle(BorderStyleType.BorderThin);
        poCtrl.setWriter(doc);

        poCtrl.setSaveFilePage("/save/common");



        poCtrl.webOpen(dir+"test44.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word44");
        return mv;
    }
    @RequestMapping(value="/word45", method= RequestMethod.GET)
    public ModelAndView showWord45(HttpServletRequest request, Map<String,Object> map){

        ModelAndView mv = new ModelAndView("/word/Word45");
        return mv;
    }
    @RequestMapping(value="/word46", method= RequestMethod.GET)
    public ModelAndView showWord46(HttpServletRequest request, Map<String,Object> map){
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/poserver.zz");
        WordDocument doc = new WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");
        fmCtrl.setSaveFilePage("/save/doc/data20");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.fillDocumentAsPDF(dir+"test46.doc", DocumentOpenType.Word, "a.pdf");




        map.put("pageoffice",fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word46");
        return mv;
    }

    @RequestMapping(value="/word47", method= RequestMethod.GET)
    public ModelAndView showWord47(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        // Create custom toolbar
        poCtrl.addCustomToolButton("保存", "SaveDocument()", 1);
        poCtrl.addCustomToolButton("显示A文档", "ShowFile1View()", 0);
        poCtrl.addCustomToolButton("显示B文档", "ShowFile2View()", 0);
        poCtrl.addCustomToolButton("显示比较结果", "ShowCompareView()", 0);


        poCtrl.wordCompare(dir+"test47//aaa1.doc", dir+"test47//aaa2.doc", OpenModeType.docAdmin, "张三");



        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word47");
        return mv;
    }

    @RequestMapping(value="/word48", method= RequestMethod.GET)
    public ModelAndView showWord48(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        doc.openDataRegion("PO_company").setValue("北京幻想科技有限公司");
        doc.openDataRegion("PO_logo").setValue("[image]/word/logo.gif[/image]");
        doc.openDataRegion("PO_dr1").setValue("左边的文本:xxxx");

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文件
        poCtrl.webOpen(dir+"test48.doc", OpenModeType.docNormalEdit,"zhangsan");



        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word48");
        return mv;
    }


    @RequestMapping(value="/word49", method= RequestMethod.GET)
    public ModelAndView showWord49(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.getRibbonBar().setTabVisible("TabHome",true);//开始
        poCtrl.getRibbonBar().setTabVisible("TabPageLayoutWord", false);//页面布局
        poCtrl.getRibbonBar().setTabVisible("TabReferences", false);//引用
        poCtrl.getRibbonBar().setTabVisible("TabMailings", false);//邮件
        poCtrl.getRibbonBar().setTabVisible("TabReviewWord", false);//审阅
        poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
        poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图


        poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮

        poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板
        //打开Word文件
        poCtrl.webOpen(dir+"test49.doc", OpenModeType.docNormalEdit,"zhangsan");



        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word49");
        return mv;
    }






    private  PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }








}