package com.example.demo.controller.seal;


import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PDFCtrl;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
public class SealPdfController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value = "/seal/pdf1", method = RequestMethod.GET)
    public ModelAndView addSealExcel1(HttpServletRequest request, Map<String, Object> map) {
        PDFCtrl pdfCtrl1 = new PDFCtrl(request);
        pdfCtrl1.setServerPage(request.getContextPath() + "/poserver.zz"); //此行必须
        //设置保存页面
        //pdfCtrl1.setSaveFilePage("/save/seal/pdf1");


        // Create custom toolbar
        pdfCtrl1.addCustomToolButton("保存", "Save()", 1);
        pdfCtrl1.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        pdfCtrl1.addCustomToolButton("加盖印章2", "InsertSeal2()", 2);
        pdfCtrl1.addCustomToolButton("删除指定印章1", "DeleteSeal1()", 21);
        pdfCtrl1.addCustomToolButton("删除指定印章3", "DeleteSeal3()", 21);
        pdfCtrl1.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        pdfCtrl1.addCustomToolButton("打印", "PrintFile()", 6);
        pdfCtrl1.addCustomToolButton("隐藏/显示书签", "SetBookmarks()", 0);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("实际大小", "SetPageReal()", 16);
        pdfCtrl1.addCustomToolButton("适合页面", "SetPageFit()", 17);
        pdfCtrl1.addCustomToolButton("适合宽度", "SetPageWidth()", 18);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("首页", "FirstPage()", 8);
        pdfCtrl1.addCustomToolButton("上一页", "PreviousPage()", 9);
        pdfCtrl1.addCustomToolButton("下一页", "NextPage()", 10);
        pdfCtrl1.addCustomToolButton("尾页", "LastPage()", 11);
        pdfCtrl1.addCustomToolButton("-", "", 0);
        pdfCtrl1.addCustomToolButton("向左旋转90度", "SetRotateLeft()", 12);
        pdfCtrl1.addCustomToolButton("向右旋转90度", "SetRotateRight()", 13);
        pdfCtrl1.webOpen(dir + "seal\\" + "pgoffice.pdf");

        map.put("pageoffice", pdfCtrl1.getHtmlCode("PDFCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/pdf/pdf1");
        return mv;
    }

    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.setZoomSealServer("http://xqx.zoomseal.cn:8080/ZoomSealEnt/enserver.zz");
        return poCtrl;
    }
}
