package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.excelreader.Sheet;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelreader.Workbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@Controller
public class SealSaveController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping("/save/seal/word1")
    public void saveSealWord1(HttpServletRequest request, HttpServletResponse response) {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);
        //保存文件到本地磁盘
        fs.saveToFile(dir + "seal\\" + fs.getFileName());
        fs.close();

    }

    @RequestMapping("/save/seal/word2")
    public ModelAndView saveSealWord2(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) {
        String filePath = dir + "seal\\" + "test11\\";
        String id = request.getParameter("id").trim();
        if ("1".equals(id)) {
            filePath = dir + "seal\\" + "test11\\" + "doc/test1.doc";
        }
        if ("2".equals(id)) {
            filePath = dir + "seal\\" + "test11\\" + "doc/test2.doc";
        }
        if ("3".equals(id)) {
            filePath = dir + "seal\\" + "test11\\" + "doc/test3.doc";
        }
        if ("4".equals(id)) {
            filePath = dir + "seal\\" + "test11\\" + "doc/test4.doc";
        }

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage(request.getContextPath() + "/poserver.zz");
        fmCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened()");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("/save/seal/word3");
        fmCtrl.fillDocument(filePath, DocumentOpenType.Word);
        map.put("pageoffice", fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/word/Word12");
        return mv;

    }


    @RequestMapping("/save/seal/excel1")
    public void saveSealExcel1(HttpServletRequest request, HttpServletResponse response) {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);
        //保存文件到本地磁盘
        fs.saveToFile(dir + "seal\\" + fs.getFileName());
        fs.close();

    }

    @RequestMapping("/save/seal/word3")
    public void saveSealWord3(HttpServletRequest request, HttpServletResponse response) {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);
        //保存文件到本地磁盘
        fs.saveToFile(dir + "seal\\" + "test11\\" + fs.getFileName());
        fs.close();

    }


}
