package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.ThemeType;
import com.zhuozhengsoft.pageoffice.excelwriter.Cell;
import com.zhuozhengsoft.pageoffice.excelwriter.Sheet;
import com.zhuozhengsoft.pageoffice.excelwriter.Workbook;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.beans.factory.annotation.Value;

import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.xml.ws.RequestWrapper;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Map;

@Controller
public class HelloController {
    @Value("${testPath}")
    private String dir;

    @Value("${posyspath}")
    private String poSysPath;
    @Value("${popassword}")
    private String poPassWord;


    @Bean
    public ServletRegistrationBean servletRegistrationBean() {
        com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
        //设置PageOffice注册成功后,license.lic文件存放的目录
        poserver.setSysPath(poSysPath);
        ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
        srb.addUrlMappings("/poserver.zz");
        srb.addUrlMappings("/posetup.exe");
        srb.addUrlMappings("/pageoffice.js");
        srb.addUrlMappings("/jquery.min.js");
        srb.addUrlMappings("/pobstyle.css");
        srb.addUrlMappings("/sealsetup.exe");
        return srb;//
    }

    @Bean
    public ServletRegistrationBean servletRegistrationBean2() {
        com.zhuozhengsoft.pageoffice.poserver.AdminSeal adminSeal = new com.zhuozhengsoft.pageoffice.poserver.AdminSeal();
        adminSeal.setAdminPassword(poPassWord);//设置印章管理员admin的登录密码
        adminSeal.setSysPath(poSysPath);//设置印章数据库文件poseal.db存放目录
        ServletRegistrationBean srb = new ServletRegistrationBean(adminSeal);
        srb.addUrlMappings("/adminseal.zz");
        srb.addUrlMappings("/sealimage.zz");
        srb.addUrlMappings("/loginseal.zz");
        return srb;//
    }


    @RequestMapping("/helloHtml")
    public String helloHtml(Map<String, Object> map) {

        map.put("hello", "from TemplateController.helloHtml");
        return "/helloHtml";
    }

    @RequestMapping(value = "/index", method = RequestMethod.GET)
    public ModelAndView showIndex() {
        ModelAndView mv = new ModelAndView("Index");
        return mv;
    }


    @RequestMapping(value = "/ppt", method = RequestMethod.GET)
    public ModelAndView showPpt(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        poCtrl.addCustomToolButton("保存", "Save", 1); //添加自定义按钮
        poCtrl.setSaveFilePage("/save");//设置保存的action
        poCtrl.addCustomToolButton("盖章", "AddSeal", 2); //添加自定义盖章按钮


        poCtrl.webOpen(dir + "test.ppt", OpenModeType.pptNormalEdit, "张三");
//        poCtrl.setCaption("项");
//        poCtrl.setTitlebar(false); //隐藏标题栏
//        poCtrl.setMenubar(false); //隐藏菜单栏
//        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
//        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        //设置打开的Word文档的主题
        //poCtrl.setTheme(ThemeType.Office2007);


        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word");
        return mv;
    }


}
