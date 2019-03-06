package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.*;
import java.util.Map;

@Controller
public class OpenFileController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value="/openWord", method= RequestMethod.GET)
    public void openWord(HttpServletRequest request, HttpServletResponse response) throws SQLException, ClassNotFoundException, IOException {

        String id = "2";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            id = request.getParameter("id");
        }
        Class.forName("org.sqlite.JDBC");

        String strUrl = "jdbc:sqlite:"+dir+"demodata\\DataBase.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from stream where id = "
                + id);
        int newID = 1;
        if (rs.next()) {
            //******读取磁盘文件，输出文件流 开始*******************************
            byte[] imageBytes = rs.getBytes("Word");
            int fileSize = imageBytes.length;

            response.reset();
            response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
            response.setHeader("Content-Disposition",
                    "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
            response.setContentLength(fileSize);

            OutputStream outputStream = response.getOutputStream();
            outputStream.write(imageBytes);

            outputStream.flush();
            outputStream.close();
            outputStream = null;
            //******读取磁盘文件，输出文件流 结束*******************************
        }
        rs.close();
        conn.close();
    }

    @RequestMapping(value="/openWord2", method= RequestMethod.GET)
    public ModelAndView openWord2(HttpServletRequest request, HttpServletResponse response, Map<String,Object> map) throws SQLException, ClassNotFoundException, IOException {

        String subject="";
        String fileName="";

        Class.forName("org.sqlite.JDBC");

        String strUrl = "jdbc:sqlite:"+dir+"demodata\\CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String id=request.getParameter("id");
        if(!id.equals("")&&!id.equals(null)){
            ResultSet rs=stmt.executeQuery("select * from word where ID="+id);
            subject=rs.getString("Subject");
            fileName=rs.getString("FileName");
            rs.close();
            System.out.println(subject);
            System.out.println(fileName);


        }
        stmt.close();
        conn.close();




        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet

        poCtrl.setAllowCopy(false);//禁止拷贝
        poCtrl.setMenubar(false);//隐藏菜单栏
        poCtrl.setOfficeToolbars(false);//隐藏Office工具条
        poCtrl.setCustomToolbar(false);//隐藏自定义工具栏
        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");
        //设置页面的显示标题

        poCtrl.setCaption("演示：文件在线安全浏览");
        //打开文件
        poCtrl.webOpen(dir+"test18.doc", OpenModeType.docNormalEdit,"张三");

        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));

        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("Word18");
        return mv;


    }



}
