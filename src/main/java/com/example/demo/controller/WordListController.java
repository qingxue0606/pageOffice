package com.example.demo.controller;

import com.example.demo.entity.Doc;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
public class WordListController {


    @RequestMapping(value="/wordlists", method= RequestMethod.GET)
    public ModelAndView showWord20(HttpServletRequest request, Map<String,Object> map) throws ClassNotFoundException, SQLException, ParseException {
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:D:\\test\\demodata\\CreateWord.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs=stmt.executeQuery("select * from word order by id desc");
        String fileName="";
        String subject="";
        String submitTime="";
        List<Doc> list=new ArrayList<>();

        while(rs.next()){
            int id=rs.getInt("ID");
            fileName = rs.getString("FileName");
            subject = rs.getString("Subject");
            submitTime = rs.getString("SubmitTime");
            if(submitTime!=null&&submitTime.length()>0){
                submitTime=new SimpleDateFormat("yyyy/MM/dd")
                        .format(new SimpleDateFormat("yyyy-MM-dd")
                                .parse(submitTime));
            }


            Doc doc=new Doc();
            doc.setId(id);
            doc.setFileName(fileName);
            doc.setSubject(subject);
            doc.setSubmitTime(submitTime);
            System.out.println(doc);
            list.add(doc);


        }
        rs.close();
        stmt.close();
        conn.close();




        map.put("list",list);



        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/list/list");


        return mv;
    }
}
