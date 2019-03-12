package com.example.demo.controller.seal;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
public class SealController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value="/seal/word1", method= RequestMethod.GET)
    public ModelAndView addSealWord1(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test1.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word1");
        return mv;
    }
    @RequestMapping(value="/seal/word2", method= RequestMethod.GET)
    public ModelAndView addSealWord2(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test2.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word2");
        return mv;
    }
    @RequestMapping(value="/seal/word3", method= RequestMethod.GET)
    public ModelAndView addSealWord3(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        //设置保存页面




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test3.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word3");
        return mv;
    }
    @RequestMapping(value="/seal/word4", method= RequestMethod.GET)
    public ModelAndView addSealWord4(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("添加印章位置", "InsertSealPos()", 2);




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test4.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word4");
        return mv;
    }
    @RequestMapping(value="/seal/word5", method= RequestMethod.GET)
    public ModelAndView addSealWord5(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test5.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word5");
        return mv;
    }
    @RequestMapping(value="/seal/word6", method= RequestMethod.GET)
    public ModelAndView addSealWord6(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test6.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word6");
        return mv;
    }
    @RequestMapping(value="/seal/word7", method= RequestMethod.GET)
    public ModelAndView addSealWord7(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("盖章到印章位置", "AddSealByPos()", 2);




        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test7.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word7");
        return mv;
    }
    @RequestMapping(value="/seal/word8", method= RequestMethod.GET)
    public ModelAndView addSealWord8(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);



        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test8.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word8");
        return mv;
    }
    @RequestMapping(value="/seal/word9", method= RequestMethod.GET)
    public ModelAndView addSealWord9(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);



        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test9.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word9");
        return mv;
    }
    @RequestMapping(value="/seal/word10", method= RequestMethod.GET)
    public ModelAndView addSealWord10(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖骑缝章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);



        poCtrl.setSaveFilePage("/save/seal/word1");



        poCtrl.webOpen(dir+"seal\\"+"test10.doc", OpenModeType.docNormalEdit,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word10");
        return mv;
    }
    @RequestMapping(value="/seal/word11", method= RequestMethod.GET)
    public ModelAndView addSealWord11(HttpServletRequest request, Map<String,Object> map){

        ModelAndView mv = new ModelAndView("seal/word/Word11");
        return mv;
    }




    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
