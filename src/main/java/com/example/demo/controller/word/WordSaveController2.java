package com.example.demo.controller.word;

import com.example.demo.util.URIEncoder;
import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordreader.DataRegion;
import com.zhuozhengsoft.pageoffice.wordreader.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class WordSaveController2 {
    @Value("${testPath}")
    private String dir;


    @RequestMapping("/save/doc/data27")
    public ModelAndView saveDocData27(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) throws IOException {
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        String id = request.getParameter("id");
        poCtrl.setSaveFilePage("/save/doc/data28?id=" + id);
        //打开Word文件
        poCtrl.webOpen("/word71?id=" + id, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word70");
        return mv;


    }

    @RequestMapping("/save/doc/data28")
    public void saveDocData28(HttpServletRequest request, HttpServletResponse response,Map<String, Object> map) throws IOException, ClassNotFoundException, SQLException {
        FileSaver fs = new FileSaver(request, response);
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id").trim();
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:"+dir+"demodata\\ExaminationPaper.db";
            Connection conn = DriverManager.getConnection(strUrl);
            String sql= "UPDATE  Stream SET Word=?  where ID=" + id ;
            PreparedStatement pstmt=null;
            pstmt= conn.prepareStatement(sql);
            pstmt.setBytes(1,fs.getFileBytes());
            //pstmt.setBinaryStream(1,fs.getFileStream(),fs.getFileSize());
            pstmt.executeUpdate();
            pstmt.close();
            conn.close();

            fs.setCustomSaveResult("ok");
        } else {
            err = "<script>alert('未获得文件的ID，保存失败');</script>";
        }
        fs.close();


    }



    private PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }



}
