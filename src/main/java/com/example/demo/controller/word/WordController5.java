package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.*;
import java.sql.*;
import java.util.Map;

@Controller
public class WordController5 {
    @Value("${testPath}")
    private String dir;


    @RequestMapping(value = "/word69", method = RequestMethod.GET)
    public ModelAndView showWord69(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException {
//--- PageOffice的调用代码 开始 -----
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + dir + "demodata\\ExaminationPaper.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("Select * from stream");
        boolean flg = false;//标识是否有数据
        StringBuilder strHtmls = new StringBuilder();
        strHtmls.append("<tr  style='background-color:#FEE;'>");
        strHtmls.append("<td style='text-align:center;width=10%' >选择</td>");
        strHtmls.append("<td style='text-align:center;width=30%'>题库编号</td>");
        strHtmls.append("<td style='text-align:center;width=60%'>操作</td>");
        strHtmls.append("</tr>");
        while (rs.next()) {
            flg = true;
            String pID = rs.getString("ID");
            strHtmls.append("<tr  style='background-color:white;'>");
            strHtmls.append("<td style='text-align:center'><input id='check" + pID + "'  type='checkbox' /></td>");
            strHtmls.append("<td style='text-align:center'>选择题-" + pID + "</td>");
            strHtmls.append("<td style='text-align:center'><a href='javascript:POBrowser.openWindowModeless(\"/save/doc/data27?id=" + pID + "\" ,\"width=1200px;height=800px;\");'>编辑</a></td>");
            strHtmls.append("</tr>");
        }

        if (!flg) {
            strHtmls.append("<tr>\r\n");
            strHtmls.append("<td width='100%' height='100' align='center'>对不起，暂时没有可以操作的数据。\r\n");
            strHtmls.append("</td></tr>\r\n");
        }


        map.put("strHtmls", strHtmls);
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word69");
        return mv;
    }


    @RequestMapping(value = "/word71", method = RequestMethod.GET)
    public void showWord71(HttpServletRequest request, Map<String, Object> map, HttpServletResponse response) throws ClassNotFoundException, SQLException, IOException {
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id");
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:" + dir + "demodata\\ExaminationPaper.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            String strSql = "select * from stream where id =" + id;
            ResultSet rs = stmt.executeQuery(strSql);
            if (rs.next()) {
                //******读取磁盘文件，输出文件流 开始*******************************
                byte[] imageBytes = rs.getBytes("Word");
                int fileSize = imageBytes.length;

                response.reset();
                response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
                response.setHeader("Content-Disposition", "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
                response.setContentLength(fileSize);

                OutputStream outputStream = response.getOutputStream();
                outputStream.write(imageBytes);

                outputStream.flush();
                outputStream.close();
                outputStream = null;
                //下面两句代码解决response.getWriter()和response.getOutputStream()冲突问题
                //out.clear();
                //out = pageContext.pushBody();

                //******读取磁盘文件，输出文件流 结束*******************************
            } else {
                err = "未获得文件的信息";
            }
            rs.close();
            stmt.close();
            conn.close();
        } else {
            err = "未获得文件的ID";
            //out.print(err);
        }
        if (err.length() > 0)
            err = "<script>alert(" + err + ");</script>";
    }


    @RequestMapping(value = "/word72", method = RequestMethod.GET)
    public ModelAndView showWord72(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);
        if (request.getParameter("ids").equals(null)
                || request.getParameter("ids").equals("")) {
        }
        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(",");//将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件即可

        int pNum = 1;
        String operateStr = "";
        operateStr += "function Create(){\n";
        // document.getElementById('PageOfficeCtrl1').Document.Application 微软office VBA对象的根Application对象
        operateStr += "var obj = document.getElementById('PageOfficeCtrl1').Document.Application;\n";
        operateStr += "obj.Selection.EndKey(6);\n"; // 定位光标到文档末尾

        for (int i = 0; i < ids.length; i++) {
            operateStr += "obj.Selection.TypeParagraph();"; //用来换行
            operateStr += "obj.Selection.Range.Text = '" + pNum + ".';\n"; // 用来生成题号
            // 下面两句代码用来移动光标位置
            operateStr += "obj.Selection.EndKey(5,1);\n";
            operateStr += "obj.Selection.MoveRight(1,1);\n";
            // 插入指定的题到文档中
            operateStr += "document.getElementById('PageOfficeCtrl1').InsertDocumentFromURL('/word71?id="
                    + ids[i] + "');\n";
            pNum++;

        }
        operateStr += "\n}\n";

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setJsFunction_AfterDocumentOpened("Create()");

        poCtrl.webOpen(dir + "test72.doc", OpenModeType.docNormalEdit, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("operateStr", operateStr);
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word72");
        return mv;
    }


    @RequestMapping(value = "/word73", method = RequestMethod.GET)
    public ModelAndView showWord73(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        String idlist = request.getParameter("ids").trim();
        String[] ids = idlist.split(","); //将idlist按照","截取后存到ids数组中，然后遍历数组用js插入文件
        String temp = "PO_begin";//存储数据区域名称
        int num = 1;//试题编号
        WordDocument doc = new WordDocument();
        for (int i = 0; i < ids.length; i++) {

            DataRegion dataNum = doc.createDataRegion("PO_" + num,
                    DataRegionInsertType.After, temp);
            dataNum.setValue(num + ".\t");
            DataRegion dataRegion = doc.createDataRegion("PO_begin"
                    + (i + 1), DataRegionInsertType.After, "PO_" + num);
            dataRegion.setValue("[word]/word71?id=" + ids[i]
                    + "[/word]");
            temp = "PO_begin" + (i + 1);
            num++;
        }


//隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.setCustomToolbar(false);
        poCtrl.setCaption("生成试卷");
        poCtrl.setWriter(doc);

        poCtrl.webOpen(dir + "test72.doc", OpenModeType.docNormalEdit, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word73");
        return mv;
    }


    @RequestMapping(value = "/word74", method = RequestMethod.GET)
    public ModelAndView showWord74(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        WordDocument doc = new WordDocument();

        //设置内容标题

        //创建DataRegion对象，PO_title为自动添加的书签名称,书签名称需以“PO_”为前缀，切书签名称不能重复
        //三个参数分别为要新插入书签的名称、新书签的插入位置、相关联的书签名称（“[home]”代表Word文档的第一个位置）
        DataRegion title = doc.createDataRegion("PO_title",
                DataRegionInsertType.After, "[home]");
        //给DataRegion对象赋值
        title.setValue("C#中Socket多线程编程实例\n");
        //设置字体：粗细、大小、字体名称、是否是斜体
        title.getFont().setBold(true);
        title.getFont().setSize(20);
        title.getFont().setName("黑体");
        title.getFont().setItalic(false);
        //定义段落对象
        ParagraphFormat titlePara = title.getParagraphFormat();
        //设置段落对齐方式
        titlePara.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        //设置段落行间距
        titlePara.setLineSpacingRule(WdLineSpacing.wdLineSpaceMultiple);

        //设置内容
        //第一段
        //创建DataRegion对象，PO_body为自动添加的书签名称
        DataRegion body = doc.createDataRegion("PO_body",
                DataRegionInsertType.After, "PO_title");
        //设置字体：粗细、是否是斜体、大小、字体名称、字体颜色
        body.getFont().setBold(false);
        body.getFont().setItalic(true);
        body.getFont().setSize(10);
        //设置中文字体名称
        body.getFont().setName("楷体");
        //设置英文字体名称
        body.getFont().setName("Times New Roman");
        body.getFont().setColor(Color.RED);
        //给DataRegion对象赋值
        body
                .setValue("是微软随着VS.net新推出的一门语言。它作为一门新兴的语言，有着C++的强健，又有着VB等的RAD特性。而且，微软推出C#主要的目的是为了对抗Sun公司的Java。大家都知道Java语言的强大功能，尤其在网络编程方面。于是，C#在网络编程方面也自然不甘落后于人。本文就向大家介绍一下C#下实现套接字（Sockets）编程的一些基本知识，以期能使大家对此有个大致了解。首先，我向大家介绍一下套接字的概念。\n");
        //创建ParagraphFormat对象
        ParagraphFormat bodyPara = body.getParagraphFormat();
        //设置段落的行间距、对齐方式、首行缩进
        bodyPara.setLineSpacingRule(WdLineSpacing.wdLineSpaceAtLeast);
        bodyPara.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara.setFirstLineIndent(0);

        //第二段
        DataRegion body2 = doc.createDataRegion("PO_body2",
                DataRegionInsertType.After, "PO_body");
        body2.getFont().setBold(false);
        body2.getFont().setSize(12);
        body2.getFont().setName("黑体");
        body2
                .setValue("套接字是通信的基石，是支持TCP/IP协议的网络通信的基本操作单元。可以将套接字看作不同主机间的进程进行双向通信的端点，它构成了单个主机内及整个网络间的编程界面。套接字存在于通信域中，通信域是为了处理一般的线程通过套接字通信而引进的一种抽象概念。套接字通常和同一个域中的套接字交换数据（数据交换也可能穿越域的界限，但这时一定要执行某种解释程序）。各种进程使用这个相同的域互相之间用Internet协议簇来进行通信。\n");
        //body2.setValue("[image]../images/logo.jpg[/image]");
        ParagraphFormat bodyPara2 = body2.getParagraphFormat();
        bodyPara2.setLineSpacingRule(WdLineSpacing.wdLineSpace1pt5);
        bodyPara2.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara2.setFirstLineIndent(21);

        //第三段
        DataRegion body3 = doc.createDataRegion("PO_body3",
                DataRegionInsertType.After, "PO_body2");
        body3.getFont().setBold(false);
        body3.getFont().setColor(Color.getHSBColor(0, 128, 228));
        body3.getFont().setSize(14);
        body3.getFont().setName("华文彩云");
        body3
                .setValue("套接字可以根据通信性质分类，这种性质对于用户是可见的。应用程序一般仅在同一类的套接字间进行通信。不过只要底层的通信协议允许，不同类型的套接字间也照样可以通信。套接字有两种不同的类型：流套接字和数据报套接字。\n");
        ParagraphFormat bodyPara3 = body3.getParagraphFormat();
        bodyPara3.setLineSpacingRule(WdLineSpacing.wdLineSpaceDouble);
        bodyPara3.setAlignment(WdParagraphAlignment.wdAlignParagraphLeft);
        bodyPara3.setFirstLineIndent(21);

        DataRegion body4 = doc.createDataRegion("PO_body4",
                DataRegionInsertType.After, "PO_body3");
        body4.setValue("[image]" + dir + "test74\\" + "logo.png[/image]");
        //body4.setValue("[word]doc/1.doc[/word]");//还可嵌入其他Word文件
        ParagraphFormat bodyPara4 = body4.getParagraphFormat();
        bodyPara4.setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

        poCtrl.setWriter(doc);
        //设置页面保存后执行的JS函数
        poCtrl.setJsFunction_AfterDocumentSaved("SaveOK()");

        //隐藏菜单栏
        poCtrl.setMenubar(false);

        poCtrl.webOpen(dir + "test74.doc", OpenModeType.docNormalEdit, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word74");
        return mv;
    }


    @RequestMapping(value = "/word75", method = RequestMethod.GET)
    public ModelAndView showWord75(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException {
//--- PageOffice的调用代码 开始 -----

        ModelAndView mv = new ModelAndView("/word/Word75");
        return mv;
    }


    @RequestMapping(value = "/word76", method = RequestMethod.GET)
    public ModelAndView showWord76(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
        poCtrl.addCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.setSaveFilePage("/save/doc/data29");

        poCtrl.webOpen(dir + "test75\\" + "test.doc", OpenModeType.docNormalEdit, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word76");
        return mv;
    }

    @RequestMapping(value = "/word77", method = RequestMethod.GET)
    public ModelAndView showWord77(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        String fileName = "";
        String mbName = request.getParameter("templateName");


        poCtrl.setCustomToolbar(false);

        if (mbName != null && mbName.trim() != "") {
            // 选择模板后执行套红

            // 复制模板，命名为正式发文的文件名：zhengshi.doc
            fileName = "zhengshi.doc";
            String templateName = request.getParameter("mb");
            String templatePath = dir + "test75\\" + templateName;
            String filePath = dir + "test75\\" + fileName;
            copyFile(templatePath, filePath);

            // 填充数据和正文内容到“zhengshi.doc”
            WordDocument doc = new WordDocument();
            DataRegion copies = doc.openDataRegion("PO_Copies");
            copies.setValue("6");
            DataRegion docNum = doc.openDataRegion("PO_DocNum");
            docNum.setValue("001");
            DataRegion issueDate = doc.openDataRegion("PO_IssueDate");
            issueDate.setValue("2013-5-30");
            DataRegion issueDept = doc.openDataRegion("PO_IssueDept");
            issueDept.setValue("开发部");
            DataRegion sTextS = doc.openDataRegion("PO_STextS");
            sTextS.setValue("[word]" + dir + "test75\\" + "test.doc[/word]");
            DataRegion sTitle = doc.openDataRegion("PO_sTitle");
            sTitle.setValue("北京某公司文件");
            DataRegion topicWords = doc.openDataRegion("PO_TopicWords");
            topicWords.setValue("Pageoffice、 套红");
            poCtrl.setWriter(doc);

        } else {
            //首次加载时，加载正文内容：test.doc
            fileName = "test.doc";

        }

        poCtrl.setSaveFilePage("/save/doc/data29");

        poCtrl.webOpen(dir + "test75\\" + fileName, OpenModeType.docNormalEdit, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word77");
        return mv;
    }


    @RequestMapping(value = "/word78", method = RequestMethod.GET)
    public ModelAndView showWord78(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = initPageOfficeCtrl(request);

        String fileName = "zhengshi.doc"; //正式发文的文件
        poCtrl.setCaption(fileName);
        poCtrl.addCustomToolButton("另存到本地", "ShowDialog1()", 5);
        poCtrl.addCustomToolButton("页面设置", "ShowDialog2()", 0);
        poCtrl.addCustomToolButton("打印", "ShowDialog3()", 6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen()", 4);

        poCtrl.setMenubar(false);
        poCtrl.setOfficeToolbars(false);

        //poCtrl.setSaveFilePage("/save/doc/data29");

        poCtrl.webOpen(dir + "test75\\" + fileName, OpenModeType.docReadOnly, "zhangsan");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word78");
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
