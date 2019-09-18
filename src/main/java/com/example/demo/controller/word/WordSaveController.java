package com.example.demo.controller.word;

import com.example.demo.util.URIEncoder;
import com.zhuozhengsoft.pageoffice.DocumentOpenType;
import com.zhuozhengsoft.pageoffice.FileMakerCtrl;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.excelreader.Sheet;
import com.zhuozhengsoft.pageoffice.excelreader.Table;
import com.zhuozhengsoft.pageoffice.excelreader.Workbook;
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
import java.nio.charset.Charset;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class WordSaveController {
    @Value("${testPath}")
    private String dir;

    @RequestMapping("/save/common")
    public void saveCommon(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + fs.getFileName());
        fs.close();

    }


    @RequestMapping("/save")
    public void saveFile(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + fs.getFileName());
        int age = 0;
        //获取通过隐藏域传递过来的值
        if (fs.getFormField("age") != null
                && fs.getFormField("age").trim().length() > 0) {
            age = Integer.parseInt(fs.getFormField("age"));
        }
        System.out.println(age);

        fs.close();

    }

    @RequestMapping("/save/doc7")
    public void saveDoc7(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + fs.getFileName());
        fs.close();

    }

    @RequestMapping("/save/doc8")
    public void saveDoc8(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + fs.getFileName());
        fs.close();

    }


    @RequestMapping("/save/doc2   ")
    public String saveDoc2(HttpServletRequest request, HttpServletResponse response) {
        WordDocument doc = new WordDocument(request, response);
        //获取提交的数值
        DataRegion dataUserName = doc.openDataRegion("PO_userName");
        byte[] bytes= dataUserName.getFileBytes();
        System.out.println(bytes.toString());


        DataRegion dataDeptName = doc.openDataRegion("PO_deptName");
        String content = "";
        //content += "公司名称：" + doc.getFormField("txtCompany");
        content += "<br/>员工姓名：" + dataUserName.getValue();
        content += "<br/>部门名称：" + dataDeptName.getValue();
        System.out.println(content);

        doc.showPage(500, 400);
        doc.close();
        return "/resp";

    }


    @RequestMapping("/save/doc/data")
    public String saveDocData(HttpServletRequest request, HttpServletResponse response) {
        WordDocument doc = new WordDocument(request, response);
        //获取提交的数值
        String dataUserName = doc.openDataRegion("PO_userName").getValue();
        String dataDeptName = doc.openDataRegion("PO_deptName").getValue();
        String companyName = doc.getFormField("txtCompany");


        doc.close();
        return "/resp";

    }

    @RequestMapping("/save/doc/data12")
    public String saveDocData12(HttpServletRequest request, HttpServletResponse response) {
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
            ErrorMsg = ErrorMsg + "<li><font color=red>注意：</font>请假天数必须是数字</li>";
        }

        if (ErrorMsg == "") {
            // 您可以在此编程，保存这些数据到数据库中。
            System.out.println("提交的数据为：<br/>");
            System.out.println("姓名：" + sName + "<br/>");
            System.out.println("部门：" + sDept + "<br/>");
            System.out.println("原因：" + sCause + "<br/>");
            System.out.println("天数：" + sNum + "<br/>");
            System.out.println("日期：" + sDate + "<br/>");
            doc.showPage(578, 380);
        } else {
            ErrorMsg = "<div style='color:#FF0000;'>请修改以下信息：</div> "
                    + ErrorMsg;
            doc.showPage(578, 380);
        }
        doc.close();
        return "/resp";

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
            pstmt.setBytes(1, fs.getFileBytes());
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
        String strUrl = "jdbc:sqlite:" + dir + "demodata\\CreateWord.db";
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
        fs.saveToFile(dir + fs.getFileName());
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
        System.out.println(fs.getFileName());
        System.out.println(fs.getFileExtName());
        fs.saveToFile(dir + fs.getFileName());
        fs.close();


    }

    @RequestMapping("/save/doc/data16")
    public ModelAndView saveDocData16(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

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
        fmCtrl.fillDocument(dir + "test28.doc", DocumentOpenType.Word);
        System.out.println(1);

        map.put("pageoffice", fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word29");
        return mv;


    }


    @RequestMapping("/save/doc/data17")
    public void saveDocData17(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

        FileSaver fs = new FileSaver(request, response);
        String fileName = "maker" + fs.getFileExtName();
        fs.saveToFile(dir + fileName);

        fs.close();

    }

    @RequestMapping("/save/doc/data18")
    public void saveDocData18(HttpServletRequest request, HttpServletResponse response) throws UnsupportedEncodingException {

        WordDocument doc = new WordDocument(request, response);
        DataRegion dataReg = doc.openDataRegion("PO_table");
        com.zhuozhengsoft.pageoffice.wordreader.Table table = dataReg.openTable(1);
        //输出提交的table中的数据

        StringBuilder dataStr = new StringBuilder();
        for (int i = 1; i <= table.getRowsCount(); i++) {

            for (int j = 1; j <= table.getColumnsCount(); j++) {
                byte[] ascii = table.openCellRC(i, j).getValue().getBytes("utf-8");
                //乱码
                String out = new String(ascii, "gb2312");
                System.out.println(out);


                dataStr.append("<div>" + table.openCellRC(i, j).getValue() + "</div>");
            }
        }

        System.out.println();


        doc.close();

    }

    @RequestMapping("/save/doc/data19")
    public void saveDocData19(HttpServletRequest request, HttpServletResponse response) throws IOException {

        WordDocument doc = new WordDocument(request, response);
        byte[] bytes = null;
        String filePath = "";
        if (request.getParameter("userName") != null && request.getParameter("userName").trim().equalsIgnoreCase("zhangsan")) {
            bytes = doc.openDataRegion("PO_com1").getFileBytes();
            filePath = "content1.doc";
        } else {
            bytes = doc.openDataRegion("PO_com2").getFileBytes();
            filePath = "content2.doc";
        }
        doc.close();

        Resource resource = new ClassPathResource("static/word/" + filePath);
        File file = resource.getFile();

        //filePath = request.getSession().getServletContext().getRealPath("SetDrByUserWord2/doc/") + "/" + filePath;
        FileOutputStream outputStream = new FileOutputStream(file);
        outputStream.write(bytes);
        outputStream.flush();
        outputStream.close();

    }


    @RequestMapping("/save/doc/data20")
    public void saveDocData20(HttpServletRequest request, HttpServletResponse response) throws IOException {

        FileSaver fs = new FileSaver(request, response);
        System.out.println(fs.getFileName());

        String fileName = "testpfd" + fs.getFileExtName();
        fs.saveToFile(dir + fileName);
        fs.close();

    }

    @RequestMapping("/save/doc/data21")
    public void saveDocData21(HttpServletRequest request, HttpServletResponse response) throws IOException {

        String filePath = dir + "test50\\";
        WordDocument doc = new WordDocument(request, response);
        byte[] bWord;

        DataRegion dr1 = doc.openDataRegion("PO_test1");


        bWord = dr1.getFileBytes();

        FileOutputStream fos1 = new FileOutputStream(filePath + "new1.doc");
        fos1.write(bWord);
        fos1.flush();
        fos1.close();

        DataRegion dr2 = doc.openDataRegion("PO_test2");
        bWord = dr2.getFileBytes();
        FileOutputStream fos2 = new FileOutputStream(filePath + "new2.doc");
        fos2.write(bWord);
        fos2.flush();
        fos2.close();

        DataRegion dr3 = doc.openDataRegion("PO_test3");
        bWord = dr3.getFileBytes();
        FileOutputStream fos3 = new FileOutputStream(filePath + "new3.doc");
        fos3.write(bWord);
        fos3.flush();
        fos3.close();

//doc.showPage(500,400);
        doc.close();

    }

    @RequestMapping("/save/doc/data22")
    public void saveDocData22(HttpServletRequest request, HttpServletResponse response) throws IOException {

        FileSaver fs = new FileSaver(request, response);
        //String aa=fs.getFileExtName();
        if (fs.getFileExtName().equals(".jpg")) {
            fs.saveToFile(dir + fs.getFileName());
        } else {
            fs.saveToFile(dir + fs.getFileName());
        }
        fs.setCustomSaveResult("ok");
        fs.close();

    }

    @RequestMapping("/save/doc/data23")
    public void saveDocData23(HttpServletRequest request, HttpServletResponse response) throws IOException {

        WordDocument doc = new WordDocument(request, response);
        DataRegion dr = doc.openDataRegion("PO_image");
        //将提取的图片保存到服务器上，图片的名称为:a.jpg
        dr.openShape(1).saveAsJPG(dir + ("test61\\") + "a.jpg");
        doc.setCustomSaveResult("保存成功,文件保存到：" + request.getSession().getServletContext().getRealPath("ExtractImage/doc/") + "\\a.jpg");
        doc.close();

    }


    @RequestMapping("/save/doc/data24")
    public ModelAndView saveDocData24(HttpServletRequest request, HttpServletResponse response, Map<String, Object> map) throws ClassNotFoundException, SQLException, UnsupportedEncodingException {

        //定义保存对象
        FileMakerCtrl fmCtrl = new FileMakerCtrl(request);
        fmCtrl.setServerPage("/poserver.zz");

        String id = request.getParameter("id");

        if (id != null && id.length() > 0) {
            com.zhuozhengsoft.pageoffice.wordwriter.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordwriter.WordDocument();
            //禁用右击事件
            doc.setDisableWindowRightClick(true);
            //给数据区域赋值，即把数据填充到模板中相应的位置
            doc.openDataRegion("PO_company").setValue("学生  " + id);
            fmCtrl.setSaveFilePage("/save/doc/data25?id=" + URIEncoder.encodeURIComponent(id));
            fmCtrl.setWriter(doc);
            fmCtrl.setJsFunction_OnProgressComplete("OnProgressComplete()");
            fmCtrl.setFileTitle("newfilename.doc");
            fmCtrl.fillDocument(dir + "test63.doc", DocumentOpenType.Word);
        }

        System.out.println(1);

        map.put("pageoffice", fmCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word64");


        return mv;


    }


    @RequestMapping("/save/doc/data25")
    public void saveDocData25(HttpServletRequest request, HttpServletResponse response) {

        FileSaver fs = new FileSaver(request, response);
        String id = request.getParameter("id");
        String err = "";
        if (id != null && id.length() > 0) {
            String fileName = "student" + id + fs.getFileExtName();
            fs.saveToFile(dir + "test63\\" + fileName);
        } else {
            err = "<script>alert('未获得文件名称');</script>";
        }
        fs.close();

    }


    @RequestMapping("/save/doc/data26")
    public void saveDocData26(HttpServletRequest request, HttpServletResponse response) throws IOException {

        FileSaver fs = new FileSaver(request, response);
        //获取通过隐藏域传递过来的值
        String fileName = "";
        if (fs.getFormField("fileName") != null
                && fs.getFormField("fileName").trim().length() > 0) {
            fileName = fs.getFormField("fileName");
        }
        System.out.println(fileName);

        byte[] bWord;
        String filePath = dir + "other\\";

        bWord = fs.getFileBytes();

        FileOutputStream fos1 = new FileOutputStream(filePath + fileName + fs.getFileExtName());
        fos1.write(bWord);
        fos1.flush();
        fos1.close();

        fs.close();

    }


}
