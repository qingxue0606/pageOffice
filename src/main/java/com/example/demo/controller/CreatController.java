package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.DocumentVersion;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.sql.*;
import java.util.Map;

@Controller
public class CreatController {

    @RequestMapping(value = "/creatWord", method = RequestMethod.GET)
    public ModelAndView creatWord(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) throws SQLException, ClassNotFoundException, IOException {


        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);

        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");

        //设置保存页面
        poCtrl.setSaveFilePage("/save/doc/data14");

        //新建Word文件，webCreateNew方法中的两个参数分别指代“操作人”和“新建Word文档的版本号”
        poCtrl.webCreateNew("张佚名", DocumentVersion.Word2003);


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word21");
        return mv;


    }
}
