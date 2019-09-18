package com.example.demo.controller.word;

import com.example.demo.entity.DocSearch;
import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.net.URLDecoder;
import java.sql.*;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

@Controller
public class WordController7 {
    @Value("${testPath}")
    private String dir;


    @RequestMapping(value = "/word83", method = RequestMethod.GET)
    public ModelAndView showWord83(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

        String key = request.getParameter("Input_KeyWord");

        String sql = "";

        if (key != null && key.trim().length() > 0) {
            sql = "select * from word  where Content like '%" + URLDecoder.decode(key, "UTF-8")
                    + "%' order by ID desc";
        } else {
            sql = "select * from word order by ID desc";
        }
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + dir + "demodata\\SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery(sql);
        List<DocSearch> docSearchs = new ArrayList<DocSearch>();

        while (rs.next()) {
            DocSearch docSearch = new DocSearch();
            docSearch.setFileName(rs.getString("FileName"));
            docSearch.setContent(rs.getString("Content"));
            docSearch.setId(rs.getInt("ID"));
            docSearchs.add(docSearch);

        }

        stmt.close();
        conn.close();
        map.put("docSearchs", docSearchs);
        map.put("key", key);


        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word83");
        return mv;
    }

    @RequestMapping(value = "/word84", method = RequestMethod.GET)
    public ModelAndView showWord84(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

        int id = Integer.parseInt(request.getParameter("id"));
        //根据id查询数据库中对应的文档名称
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + dir + "demodata\\SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String sql = "select * from word where id=" + id;

        ResultSet rs = stmt.executeQuery(sql);
        String FileName = "";
        while (rs.next()) {
            FileName = rs.getString("FileName");
        }
        stmt.close();
        conn.close();

        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("/save/doc/data32?id=" + id);

        //打开Word文件
        String filePath = dir + "test83\\" + FileName + ".doc";

        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));


        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word84");
        return mv;
    }


    @RequestMapping(value = "/word85", method = RequestMethod.GET)
    public ModelAndView showWord85(HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("/word/Word85");
        return mv;
    }

    /**
     * 批量转pdf
     *
     * @param request
     * @param map
     * @return
     */
    @RequestMapping(value = "/word86", method = RequestMethod.GET)
    public ModelAndView showWord86(HttpServletRequest request, Map<String, Object> map) {
        String filePath = dir + "test85\\";
        String id = request.getParameter("id").trim();

        if ("1".equals(id)) {
            filePath = filePath + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = filePath + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = filePath + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = filePath + "PageOffice产品对客户端环境要求.doc";
        }

        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/poserver.zz");
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setSaveFilePage("/save/doc/data33");
        fmCtrl.fillDocumentAsPDF(filePath, DocumentOpenType.Word, "a.pdf");

        map.put("pageoffice", fmCtrl.getHtmlCode("PageOfficeCtrl1"));


        ModelAndView mv = new ModelAndView("/word/Word86");
        return mv;
    }


    @RequestMapping(value = "/word87", method = RequestMethod.GET)
    public ModelAndView showWord87(HttpServletRequest request, Map<String, Object> map) {
        String filePath = dir + "test85\\";
        String id = request.getParameter("id").trim();

        if ("1".equals(id)) {
            filePath = filePath + "PageOffice产品简介.doc";
        }
        if ("2".equals(id)) {
            filePath = filePath + "Pageoffice客户端安装步骤.doc";
        }
        if ("3".equals(id)) {
            filePath = filePath + "PageOffice的应用领域.doc";
        }
        if ("4".equals(id)) {
            filePath = filePath + "PageOffice产品对客户端环境要求.doc";
        }

        PageOfficeCtrl poCtrl1 = initPageOfficeCtrl(request);
        poCtrl1.setSaveFilePage("/save/doc/data33");//如要保存文件，此行必须
        poCtrl1.addCustomToolButton("保存", "Save()", 1);//添加自定义工具栏按钮
        poCtrl1.webOpen(filePath, OpenModeType.docNormalEdit, "张三");

        map.put("pageoffice", poCtrl1.getHtmlCode("PageOfficeCtrl1"));


        ModelAndView mv = new ModelAndView("/word/Word87");
        return mv;
    }


    @RequestMapping(value = "/word88", method = RequestMethod.GET)
    public ModelAndView showWord88(HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("/word/Word88");
        return mv;
    }


    @RequestMapping(value = "/word89", method = RequestMethod.GET)
    public ModelAndView showWord89(HttpServletRequest request, Map<String, Object> map) {

        PageOfficeCtrl poCtrl1 = initPageOfficeCtrl(request);
        //添加自定义按钮
        poCtrl1.addCustomToolButton("关闭","Close",21);


        poCtrl1.webOpen(dir+"test89.doc", OpenModeType.docNormalEdit, "张三");

        map.put("pageoffice", poCtrl1.getHtmlCode("PageOfficeCtrl1"));


        ModelAndView mv = new ModelAndView("/word/Word89");
        return mv;
    }





    // 拷贝文件
    private void copyFile(String oldPath, String newPath) {
        try {
            int bytesum = 0;
            int byteread = 0;
            File oldfile = new File(oldPath);
            if (oldfile.exists()) { //文件存在时
                InputStream inStream = new FileInputStream(oldPath); //读入原文件
                FileOutputStream fs = new FileOutputStream(newPath);
                byte[] buffer = new byte[1444];
                int length;
                while ((byteread = inStream.read(buffer)) != -1) {
                    bytesum += byteread; //字节数 文件大小
                    //System.out.println(bytesum);
                    fs.write(buffer, 0, byteread);
                }
                inStream.close();
            }
        } catch (Exception e) {
            System.out.println("复制单个文件操作出错");
            e.printStackTrace();
        }
    }


    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }


}
