package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
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
import java.io.UnsupportedEncodingException;
import java.util.Map;

@Controller
public class TestWordController2 {
    @Value("${testPath}")
    private String dir;


    @RequestMapping(value = "/xiang/iframOut", method = RequestMethod.GET)
    public String showIframOut(HttpSession session){


        return "xiang/iframOut";

    }
    @RequestMapping(value = "/xiang/iframIn", method = RequestMethod.GET)
    public String showIframIn(HttpSession session){


        return "xiang/iframIn";

    }







}
