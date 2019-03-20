package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
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
import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class TestWordController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value="/xiang/word", method= RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String,Object> map){
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存","Save",1); //添加自定义按钮
        poCtrl.setSaveFilePage("/test/save/doc1");//设置保存的action
        poCtrl.addCustomToolButton("盖章","AddSeal",2); //添加自定义盖章按钮
        poCtrl.addCustomToolButton("自定义","Test",3); //添加自定义盖章按钮


        //新建一个WordDocument用来操作数据
        WordDocument doc = new WordDocument();

        //打开数据区域
        DataRegion dataRegion1 = doc.openDataRegion("PO_userName");
        //给数据区域赋值
        dataRegion1.setValue("张三");

        DataRegion dataRegion2 = doc.openDataRegion("PO_deptName");
        dataRegion2.setValue("销售部");

        poCtrl.setWriter(doc);


        //poCtrl.webOpen(dir+"xiang\\"+"2974-居住证明.docx", OpenModeType.docAdmin,"张三");
        poCtrl.webOpen("/1.doc", OpenModeType.docAdmin,"张三");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("xiang/Word");
        return mv;
    }
    @RequestMapping("/test/save/doc1")
    public void saveDoc8(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir +"xiang\\"+ fs.getFileName());
        fs.close();

    }






    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
