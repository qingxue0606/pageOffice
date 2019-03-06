package com.example.demo.controller;

import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.excelreader.Sheet;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelreader.Workbook;
import com.zhuozhengsoft.pageoffice.wordreader.DataRegion;
import com.zhuozhengsoft.pageoffice.wordreader.WordDocument;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class SaveController {
    @Value("d:\\test\\")
    private String dir;

    @RequestMapping("/save")
    public void saveFile(HttpServletRequest request, HttpServletResponse response){
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile("d:\\test\\" + fs.getFileName());
        int age=0;
        //获取通过隐藏域传递过来的值
        if (fs.getFormField("age") != null
                && fs.getFormField("age").trim().length() > 0) {
            age = Integer.parseInt(fs.getFormField("age"));
        }
        System.out.println(age);

        fs.close();

    }

    @RequestMapping("/save/doc7")
    public void saveDoc7(HttpServletRequest request, HttpServletResponse response){
        FileSaver fs=new FileSaver(request,response);
        fs.saveToFile("d:\\test\\" + fs.getFileName());
        fs.close();

    }

    @RequestMapping("/save/doc8")
    public void saveDoc8(HttpServletRequest request, HttpServletResponse response){
        FileSaver fs=new FileSaver(request,response);
        fs.saveToFile("d:\\test\\" + fs.getFileName());
        fs.close();

    }


    @RequestMapping("/save/doc2")
    public String saveDoc2(HttpServletRequest request, HttpServletResponse response){
        WordDocument doc = new WordDocument(request, response);
        //获取提交的数值
        DataRegion dataUserName = doc.openDataRegion("PO_userName");
        DataRegion dataDeptName = doc.openDataRegion("PO_deptName");
        String content = "";
        //content += "公司名称：" + doc.getFormField("txtCompany");
        content += "<br/>员工姓名：" + dataUserName.getValue();
        content += "<br/>部门名称：" + dataDeptName.getValue();
        System.out.println(content);

        doc.showPage(500, 400);
        doc.close();
        return"/resp";

    }


    @RequestMapping("/save/exl2")
    public String saveExl2(HttpServletRequest request, HttpServletResponse response){
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("Info");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                //out.print(table.getDataFields().get(2).getText()+"      mmmmmmmmmmmmm          "+table.getDataFields().get(1).getText());
                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length()==0
                ) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df=(DecimalFormat) NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f*100)+"%";
                }
                content += "<br/>*********************************************";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();

        workBook.showPage(500, 400);
        workBook.close();
        return"/resp";

    }

    @RequestMapping("/save/doc/data")
    public String saveDocData(HttpServletRequest request, HttpServletResponse response){
        WordDocument doc = new WordDocument(request, response);
        //获取提交的数值
        String dataUserName = doc.openDataRegion("PO_userName").getValue();
        String dataDeptName = doc.openDataRegion("PO_deptName").getValue();
        String companyName= doc.getFormField("txtCompany");


        doc.close();
        return"/resp";

    }

    @RequestMapping("/save/doc/data12")
    public String saveDocData12(HttpServletRequest request, HttpServletResponse response){
        String ErrorMsg = "";
        String BaseUrl = "";
        //-----------  PageOffice 服务器端编程开始  -------------------//
        WordDocument doc = new WordDocument(request, response);
        String sName = doc.openDataRegion("PO_name").getValue();
        String sDept = doc.openDataRegion("PO_dept").getValue();
        String sCause = doc.openDataRegion("PO_cause").getValue();
        String sNum = doc.openDataRegion("PO_num").getValue();
        String sDate = doc.openDataRegion("PO_date").getValue();

        if (sName.equals("")) {
            ErrorMsg = ErrorMsg + "<li>申请人</li>";
        }
        if (sDept.equals("")) {
            ErrorMsg = ErrorMsg + "<li>部门名称</li>";
        }
        if (sCause.equals("")) {
            ErrorMsg = ErrorMsg + "<li>请假原因</li>";
        }
        if (sDate.equals("")) {
            ErrorMsg = ErrorMsg + "<li>日期</li>";
        }
        try {
            if (sNum != "") {
                if (Integer.parseInt(sNum) < 0) {
                    ErrorMsg = ErrorMsg + "<li>请假天数不能是负数</li>";
                }
            } else {
                ErrorMsg = ErrorMsg + "<li>请假天数</li>";
            }
        } catch (Exception Ex) {
            ErrorMsg = ErrorMsg	+ "<li><font color=red>注意：</font>请假天数必须是数字</li>";
        }

        if (ErrorMsg == "") {
            // 您可以在此编程，保存这些数据到数据库中。
            System.out.println("提交的数据为：<br/>");
            System.out.println("姓名："+sName+"<br/>");
            System.out.println("部门："+sDept+"<br/>");
            System.out.println("原因："+sCause+"<br/>");
            System.out.println("天数："+sNum+"<br/>");
            System.out.println("日期："+sDate+"<br/>");
            doc.showPage(578, 380);
        } else {
            ErrorMsg = "<div style='color:#FF0000;'>请修改以下信息：</div> "
                    + ErrorMsg;
            doc.showPage(578, 380);
        }
        doc.close();
        return"/resp";

    }

    @RequestMapping("/save/doc/data13")
    public void saveDocData13(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, SQLException {
        FileSaver fs = new FileSaver(request, response);
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id").trim();
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:D:\\test\\demodata\\DataBase.db";
            Connection conn = DriverManager.getConnection(strUrl);
            String sql = "UPDATE  Stream SET Word=?  where ID=" + id;
            PreparedStatement pstmt = null;
            pstmt = conn.prepareStatement(sql);
            pstmt.setBytes(1,fs.getFileBytes());
            //pstmt.setBinaryStream(1, fs.getFileStream(),fs.getFileSize());
            pstmt.executeUpdate();
            pstmt.close();
            conn.close();

            fs.setCustomSaveResult("ok");
        } else {
            err = "<script>alert('未获得文件的ID，保存失败');</script>";
        }
        fs.close();

    }



    @RequestMapping("/save/doc/data14")
    public void saveDocData14(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);

        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:D:\\test\\demodata\\CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select Max(ID) from word");
        int newID = 1;
        if (rs.next()) {
            newID = Integer.parseInt(rs.getString(1)) + 1;
        }
        rs.close();

        String FileSubject = fs.getFormField("FileSubject").trim();
        String fileName = "aabb" + newID + ".doc";
        String getFile = (String) request.getParameter("FileSubject");
        if (getFile != null && getFile.length() > 0)
            FileSubject = new String(getFile.getBytes("iso-8859-1"));
        //out.print(FileSubject);
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        // new Date()为获取当前系统时间
        String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values("
                + newID
                + ",'"
                + fileName
                + "','"
                + FileSubject
                + "','"
                + df.format(new Date()) + "')";
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();

        //保存文件
        fs.saveToFile("d:\\test\\" + fs.getFileName());
        //设置保存结果
        fs.setCustomSaveResult("ok");
        //fs.showPage(300,300);
        fs.close();

    }

    @RequestMapping("/save/doc/data15")
    public void saveDocData15(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {
        //定义保存对象
        FileSaver fs = new FileSaver(request, response);
        //保存文件到本地磁盘
        fs.saveToFile(dir+fs.getFileName());
        fs.close();



    }

    @RequestMapping("/save/doc/data16")
    public ModelAndView saveDocData16(HttpServletRequest request, HttpServletResponse response,Map<String,Object> map) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

        //定义保存对象
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);

        fmCtrl.setServerPage("/poserver.zz");
        com.zhuozhengsoft.pageoffice.wordwriter.WordDocument doc =
                new com.zhuozhengsoft.pageoffice.wordwriter.WordDocument();
        //禁用右击事件
        doc.setDisableWindowRightClick(true);
        //给数据区域赋值，即把数据填充到模板中相应的位置
        doc.openDataRegion("PO_company").setValue("北京卓正志远软件有限公司  ");

        fmCtrl.setSaveFilePage("/save/doc/data17");
        fmCtrl.setWriter(doc);
        fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
        fmCtrl.setFileTitle("newfilename.doc");
        fmCtrl.fillDocument(dir+"test28.doc", DocumentOpenType.Word);
        System.out.println(1);

        map.put("pageoffice",fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word29");
        return mv;


    }


    @RequestMapping("/save/doc/data17")
    public void saveDocData17(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {
        System.out.println(2);
        FileSaver fs = new FileSaver(request, response);
        String fileName = "maker" + fs.getFileExtName();
        fs.saveToFile(dir+ fileName);
        fs.close();



    }






    @RequestMapping("/save/exl/data1")
    public String saveExlData1(HttpServletRequest request, HttpServletResponse response){
        Workbook workBook = new Workbook(request, response);
        Sheet sheet = workBook.openSheet("Sheet1");
        Table table = sheet.openTable("B4:F13");
        String content = "";
        int result = 0;
        while (!table.getEOF()) {
            //获取提交的数值
            if (!table.getDataFields().getIsEmpty()) {
                content += "<br/>月份名称："
                        + table.getDataFields().get(0).getText();
                content += "<br/>计划完成量："
                        + table.getDataFields().get(1).getText();
                content += "<br/>实际完成量："
                        + table.getDataFields().get(2).getText();
                content += "<br/>累计完成量："
                        + table.getDataFields().get(3).getText();
                if (table.getDataFields().get(2).getText().equals(null)
                        || table.getDataFields().get(2).getText().trim().length()==0
                ) {
                    content += "<br/>完成率：0%";
                } else {
                    float f = Float.parseFloat(table.getDataFields().get(2)
                            .getText());
                    f = f / Float.parseFloat(table.getDataFields().get(1).getText());
                    DecimalFormat df=(DecimalFormat)NumberFormat.getInstance();
                    content += "<br/>完成率：" + df.format(f*100)+"%";
                }
                content +="</br>";
            }
            //循环进入下一行
            table.nextRow();
        }
        table.close();
        workBook.showPage(500, 400);
        workBook.close();
        System.out.println(content);
        request.setAttribute("content",content);
        return"/resp";

    }






    /*@RequestMapping("/helloHtml")
    public String helloHtml(Map<String,Object> map){

        map.put("hello","from TemplateController.helloHtml");
        return"/helloHtml";
    }*/


}
