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
public class WordController3 {
    @Value("${testPath}")
    private String dir;

    //@RequestMapping(value="/word23", method= RequestMethod.GET)
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








}
