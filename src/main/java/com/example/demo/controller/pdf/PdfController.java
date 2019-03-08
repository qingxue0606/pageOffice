package com.example.demo.controller.pdf;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@Controller
public class PdfController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value="/pdf1", method= RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
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
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.webOpen(dir+"test1.pdf");



        map.put("pageoffice",poCtrl1.getHtmlCode("PDFCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("pdf/pdf1");
        return mv;
    }



   @RequestMapping(value="/pdf2", method= RequestMethod.GET)
    public ModelAndView showPdf2(HttpServletRequest request, Map<String,Object> map){
       PDFCtrl poCtrl1 = new PDFCtrl(request);
       poCtrl1.setServerPage("/poserver.zz"); //此行必须

// Create custom toolbar
       poCtrl1.addCustomToolButton("搜索", "SearchText()", 0);
       poCtrl1.addCustomToolButton("搜索下一个", "SearchTextNext()", 0);
       poCtrl1.addCustomToolButton("搜索上一个", "SearchTextPrev()", 0);
       poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
       poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
       poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
       poCtrl1.webOpen(dir+"test2.pdf");


        map.put("pageoffice",poCtrl1.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/pdf/pdf2");
        return mv;
    }

    @RequestMapping(value="/pdf3", method= RequestMethod.GET)
    public ModelAndView showPdf3(HttpServletRequest request, Map<String,Object> map){
        PDFCtrl poCtrl1 = new PDFCtrl(request);
        poCtrl1.setServerPage("/poserver.zz"); //此行必须

// Create custom toolbar
        poCtrl1.addCustomToolButton("打印", "Print()", 6);

        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        poCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        poCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("左转", "RotateLeft()", 12);
        poCtrl1.addCustomToolButton("右转", "RotateRight()", 13);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("放大", "ZoomIn()", 14);
        poCtrl1.addCustomToolButton("缩小", "ZoomOut()", 15);
        poCtrl1.addCustomToolButton("-", "", 0);
        poCtrl1.addCustomToolButton("全屏", "SwitchFullScreen()", 4);
        poCtrl1.webOpen(dir+"test3.jpg");


        map.put("pageoffice",poCtrl1.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/pdf/pdf3");
        return mv;
    }











}
