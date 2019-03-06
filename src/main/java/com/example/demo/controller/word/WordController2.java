package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
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
public class WordController2 {
    @Value("d:\\test\\")
    private String dir;

    @RequestMapping(value="/word23", method= RequestMethod.GET)
    public ModelAndView showWord23(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument worddoc = new WordDocument();
        //先在要插入word文件的位置手动插入书签,书签必须以“PO_”为前缀
        //给DataRegion赋值,值的形式为："[word]word文件路径[/word]、[excel]excel文件路径[/excel]、[image]图片路径[/image]"
        DataRegion data1 = worddoc.openDataRegion("PO_p1");
        data1.setValue("[word]word/1.doc[/word]");
        DataRegion data2 = worddoc.openDataRegion("PO_p2");
        data2.setValue("[image]images/img_6.jpg[/image]");
        DataRegion data3 = worddoc.openDataRegion("PO_p3");
        data3.setValue("[excel]excel/1.xls[/excel]");

        poCtrl.setWriter(worddoc);
        poCtrl.setCaption("演示：后台编程插入Word文件到数据区域");

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir+"test23\\test.doc", OpenModeType.docAdmin,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word23");
        return mv;
    }

    @RequestMapping(value="/word24", method= RequestMethod.GET)
    public ModelAndView showWord24(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc =new WordDocument();
        //添加水印 ，设置水印的内容
        doc.getWaterMark().setText("PageOffice开发平台");

        poCtrl.setWriter(doc);

        poCtrl.webOpen(dir+"test24.doc", OpenModeType.docAdmin,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }


    @RequestMapping(value="/word25", method= RequestMethod.GET)
    public ModelAndView showWord25(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //定义WordDocument对象
        WordDocument doc = new WordDocument();

        //定义DataTag对象
        DataTag deptTag = doc.openDataTag("{部门名}");
        //给DataTag对象赋值
        deptTag.setValue("B部门");
        deptTag.getFont().setColor(Color.GREEN);

        DataTag userTag = doc.openDataTag("{姓名}");
        userTag.setValue("李四");
        userTag.getFont().setColor(Color.GREEN);

        DataTag dateTag = doc.openDataTag("【时间】");
        dateTag.setValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()).toString());
        dateTag.getFont().setColor(Color.BLUE);

        poCtrl.setWriter(doc);



        poCtrl.webOpen(dir+"test25.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }

    @RequestMapping(value="/word26", method= RequestMethod.GET)
    public ModelAndView showWord26(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc=new WordDocument();
        //创建数据区域，createDataRegion 方法中的三个参数分别代表“新建的数据区域名称”，“数据区域将要插入的位置”，
        //“与新建的数据区域相关联的数据区域名称”，若当前Word文档中尚无数据区域（书签）或者想在文档的最开头创建时，那么第三个参数为“[home]”
        //若想在文档的结尾处创建数据区域则第三个参数为“[end]”
        DataRegion dataRegion1 =  doc.createDataRegion("reg1",DataRegionInsertType.After,"[home]");
        //设置创建的数据区域的可编辑性
        dataRegion1.setEditing(true);
        //给数据区域赋值
        dataRegion1.setValue("第一个数据区域\r\n");

        DataRegion dataRegion2 = doc.createDataRegion("reg2", DataRegionInsertType.After,"reg1");
        dataRegion2.setEditing(true);
        dataRegion2.setValue("第二个数据区域");

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir+"test26.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }


    @RequestMapping(value="/word27", method= RequestMethod.GET)
    public ModelAndView showWord27(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir+"test27.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word27");
        return mv;
    }

    @RequestMapping(value="/word28", method= RequestMethod.GET)
    public ModelAndView showWord28(HttpServletRequest request, Map<String,Object> map){

        ModelAndView mv = new ModelAndView("/word/Word28");
        return mv;
    }





}
