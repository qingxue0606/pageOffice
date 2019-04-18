package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.*;
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
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class TestWordController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value = "/xiang/word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpSession session, HttpServletRequest request, Map<String, Object> map) {
        Object user=session.getAttribute("user");
        System.out.println(user);
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/test/save/doc1");//设置保存的action
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 3); //添加自定义盖章按钮

        poCtrl.addCustomToolButton("自定义", "Test", 4); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 5); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 6); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 7); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义", "Test", 8); //添加自定义盖章按钮



        //新建一个WordDocument用来操作数据
        WordDocument doc = new WordDocument();

        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("张三");
        dataRegion1.setEditing(false);

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");
        doc.getWaterMark().setText("xiang");

        poCtrl.setWriter(doc);


        //poCtrl.setOfficeVendor(OfficeVendorType.WPSOffice);
        poCtrl.webOpen(dir+"xiang\\"+"test1.doc", OpenModeType.docSubmitForm,"张三");
        //poCtrl.webOpen("/1.doc", OpenModeType.docAdmin, "张三");
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



        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument(dir+"xiang\\"+"test3.doc", DocumentOpenType.Word);

        map.put("pageoffice",fmCtrl.getHtmlCode("PageOfficeCtrl1"));


        ModelAndView mv = new ModelAndView("xiang/Word3");
        return mv;
    }










    @RequestMapping("/test/save/doc1")
    public void saveDoc1(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.close();

    }

    /**
     * 保存后台文件
     */

    @RequestMapping("/test/save/doc2")
    public void saveDoc2(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "xiang\\" + fs.getFileName());
        fs.close();

    }





    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
