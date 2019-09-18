package com.example.demo.controller.seal;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@Controller
public class SealExcelController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value = "/seal/excel1", method = RequestMethod.GET)
    public ModelAndView addSealExcel1(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");


        poCtrl.webOpen(dir + "seal\\" + "test1.xls", OpenModeType.xlsNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel1");
        return mv;
    }

    @RequestMapping(value = "/seal/excel2", method = RequestMethod.GET)
    public ModelAndView addSealExcel2(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");


        poCtrl.webOpen(dir + "seal\\" + "test2.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel2");
        return mv;
    }

    @RequestMapping(value = "/seal/excel3", method = RequestMethod.GET)
    public ModelAndView addSealExcel3(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test3.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel3");
        return mv;
    }

    @RequestMapping(value = "/seal/excel4", method = RequestMethod.GET)
    public ModelAndView addSealExcel4(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("验证文档", "VerifySeal()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test4.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel4");
        return mv;
    }

    @RequestMapping(value = "/seal/excel5", method = RequestMethod.GET)
    public ModelAndView addSealExcel5(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("加盖印章", "InsertSeal()", 2);
        poCtrl.addCustomToolButton("删除指定印章", "DeleteSeal()", 21);
        poCtrl.addCustomToolButton("清除所有印章", "DeleteAllSeal()", 21);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test5.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel5");
        return mv;
    }

    @RequestMapping(value = "/seal/excel6", method = RequestMethod.GET)
    public ModelAndView addSealExcel6(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        poCtrl.addCustomToolButton("删除签字", "DeleteHandSign()", 21);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test6.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel6");
        return mv;
    }

    @RequestMapping(value = "/seal/excel7", method = RequestMethod.GET)
    public ModelAndView addSealExcel7(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "InsertHandSign()", 3);
        poCtrl.addCustomToolButton("修改密码", "ChangePsw()", 0);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test7.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel7");
        return mv;
    }

    @RequestMapping(value = "/seal/excel8", method = RequestMethod.GET)
    public ModelAndView addSealExcel8(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("签字", "AddHandSign()", 3);
        poCtrl.addCustomToolButton("验证印章", "VerifySeal()", 5);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/seal/excel1");

        poCtrl.webOpen(dir + "seal\\" + "test8.xls", OpenModeType.xlsNormalEdit, "李志");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("seal/excel/Excel8");
        return mv;
    }


    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.setZoomSealServer("http://xqx.zoomseal.cn:8080/ZoomSealEnt/enserver.zz");
        return poCtrl;
    }


}
