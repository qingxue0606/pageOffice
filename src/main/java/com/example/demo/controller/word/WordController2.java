package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@Controller
public class WordController2 {
    @Value("${testPath}")
    private String dir;

    @RequestMapping(value = "/word23", method = RequestMethod.GET)
    public ModelAndView showWord23(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument worddoc = new WordDocument();
        //先在要插入word文件的位置手动插入书签,书签必须以“PO_”为前缀
        //给DataRegion赋值,值的形式为："[word]word文件路径[/word]、[excel]excel文件路径[/excel]、[image]图片路径[/image]"
        DataRegion data1 = worddoc.openDataRegion("PO_p1");
        data1.setValue("[word]word/data.doc[/word]");
        DataRegion data2 = worddoc.openDataRegion("PO_p2");
        data2.setValue("[image]images/img_6.jpg[/image]");
        DataRegion data3 = worddoc.openDataRegion("PO_p3");
        data3.setValue("[excel]excel/1.xls[/excel]");

        poCtrl.setWriter(worddoc);
        poCtrl.setCaption("演示：后台编程插入Word文件到数据区域");

        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);
        poCtrl.setSaveFilePage("/save/common");


        poCtrl.webOpen(dir + "test23\\test.doc", OpenModeType.docAdmin, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word23");
        return mv;
    }

    @RequestMapping(value = "/word24", method = RequestMethod.GET)
    public ModelAndView showWord24(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        //添加水印 ，设置水印的内容
        doc.getWaterMark().setText("PageOffice开发平台");

        poCtrl.setWriter(doc);

        poCtrl.webOpen(dir + "test24.doc", OpenModeType.docAdmin, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }


    @RequestMapping(value = "/word25", method = RequestMethod.GET)
    public ModelAndView showWord25(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //定义WordDocument对象
        WordDocument doc = new WordDocument();

        //定义DataTag对象
        DataTag deptTag = doc.openDataTag("{部门名}");
        //给DataTag对象赋值
        deptTag.setValue("B部门");
        deptTag.getFont().setColor(Color.GREEN);

        DataTag userTag = doc.openDataTag("{姓名}");
        userTag.setValue("李四");
        userTag.getFont().setColor(Color.GREEN);

        DataTag dateTag = doc.openDataTag("【时间】");
        dateTag.setValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()).toString());
        dateTag.getFont().setColor(Color.BLUE);

        poCtrl.setWriter(doc);


        poCtrl.webOpen(dir + "test25.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }

    @RequestMapping(value = "/word26", method = RequestMethod.GET)
    public ModelAndView showWord26(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        //创建数据区域，createDataRegion 方法中的三个参数分别代表“新建的数据区域名称”，“数据区域将要插入的位置”，
        //“与新建的数据区域相关联的数据区域名称”，若当前Word文档中尚无数据区域（书签）或者想在文档的最开头创建时，那么第三个参数为“[home]”
        //若想在文档的结尾处创建数据区域则第三个参数为“[end]”
        DataRegion dataRegion1 = doc.createDataRegion("reg1", DataRegionInsertType.After, "[home]");
        //设置创建的数据区域的可编辑性
        dataRegion1.setEditing(true);
        //给数据区域赋值
        dataRegion1.setValue("第一个数据区域\r\n");

        DataRegion dataRegion2 = doc.createDataRegion("reg2", DataRegionInsertType.After, "reg1");
        dataRegion2.setEditing(true);
        dataRegion2.setValue("第二个数据区域");

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir + "test26.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word24");
        return mv;
    }


    @RequestMapping(value = "/word27", method = RequestMethod.GET)
    public ModelAndView showWord27(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);


        poCtrl.webOpen(dir + "test27.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word27");
        return mv;
    }

    @RequestMapping(value = "/word28", method = RequestMethod.GET)
    public ModelAndView showWord28(HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("/word/Word28");
        return mv;
    }

    @RequestMapping(value = "/word30", method = RequestMethod.GET)
    public ModelAndView showWord30(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        Table table1 = doc.openDataRegion("PO_T001").openTable(1);

        table1.openCellRC(1, 1).setValue("PageOffice组件");
        int dataRowCount = 5;//需要插入数据的行数
        int oldRowCount = 3;//表格中原有的行数
        // 扩充表格
        for (int j = 0; j < dataRowCount - oldRowCount; j++) {
            table1.insertRowAfter(table1.openCellRC(2, 5));  //在第2行的最后一个单元格下插入新行
        }
        // 填充数据
        int i = 1;
        while (i <= dataRowCount) {
            table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
            table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
            table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
            table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
            i++;
        }
        table1.openColumn(1).setWidth(50.2f);
        table1.openColumn(2).setWidth(100.2f);
        table1.openColumn(3).setWidth(150.2f);
        table1.openColumn(4).setWidth(200.2f);
        poCtrl.setWriter(doc);


        poCtrl.webOpen(dir + "test30.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word30");
        return mv;
    }

    @RequestMapping(value = "/word31", method = RequestMethod.GET)
    public ModelAndView showWord31(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.addCustomToolButton("开始手写", "StartHandDraw()", 5);
        poCtrl.addCustomToolButton("设置线宽", "SetPenWidth()", 5);
        poCtrl.addCustomToolButton("设置颜色", "SetPenColor()", 5);
        poCtrl.addCustomToolButton("设置笔型", "SetPenType()", 5);
        poCtrl.addCustomToolButton("设置缩放", "SetPenZoom()", 5);
        poCtrl.addCustomToolButton("访问手写集", "GetHandDrawList()", 6);

        poCtrl.setSaveFilePage("/save/common");


        poCtrl.webOpen(dir + "test31.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word31");
        return mv;
    }

    @RequestMapping(value = "/word32", method = RequestMethod.GET)
    public ModelAndView showWord32(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dTable = doc.openDataRegion("PO_table");
        //设置数据区域可编辑性
        dTable.setEditing(true);

        //打开数据区域中的表格，OpenTable(index)方法中的index为word文档中表格的下标，从1开始
        Table table1 = doc.openDataRegion("PO_Table").openTable(1);
        //设置表格边框样式
        table1.getBorder().setLineColor(Color.green);
        table1.getBorder().setLineWidth(WdLineWidth.wdLineWidth050pt);
        // 设置表头单元格文本居中
        table1.openCellRC(1, 2).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(1, 3).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(2, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(3, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

        // 给表头单元格赋值
        table1.openCellRC(1, 2).setValue("产品1");
        table1.openCellRC(1, 3).setValue("产品2");
        table1.openCellRC(2, 1).setValue("A部门");
        table1.openCellRC(3, 1).setValue("B部门");

        poCtrl.setWriter(doc);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);

        poCtrl.setSaveDataPage("/save/doc/data18");


        poCtrl.webOpen(dir + "test32.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word32");
        return mv;
    }

    @RequestMapping(value = "/word33", method = RequestMethod.GET)
    public ModelAndView showWord33(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        WordDocument doc = new WordDocument();
        DataRegion d1 = doc.openDataRegion("d1");
        d1.getFont().setColor(Color.BLUE);//设置数据区域文本字体颜色
        d1.getFont().setName("华文彩云");//设置数据区域文本字体样式
        d1.getFont().setSize(16);//设置数据区域文本字体大小
        d1.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphCenter);//设置数据区域文本对齐方式

        DataRegion d2 = doc.openDataRegion("d2");
        d2.getFont().setColor(Color.orange);//设置数据区域文本字体颜色
        d2.getFont().setName("黑体");//设置数据区域文本字体样式
        d2.getFont().setSize(14);//设置数据区域文本字体大小
        d2.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphLeft);//设置数据区域文本对齐方式

        DataRegion d3 = doc.openDataRegion("d3");
        d3.getFont().setColor(Color.magenta);//设置数据区域文本字体颜色
        d3.getFont().setName("华文行楷");//设置数据区域文本字体样式
        d3.getFont().setSize(12);//设置数据区域文本字体大小
        d3.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphRight);//设置数据区域文本对齐方式
        poCtrl.setWriter(doc);


        poCtrl.webOpen(dir + "test33.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word33");
        return mv;
    }

    @RequestMapping(value = "/word34", method = RequestMethod.GET)
    public ModelAndView showWord34(HttpServletRequest request, Map<String, Object> map) {


        ModelAndView mv = new ModelAndView("/word/Word34");
        return mv;
    }

    @RequestMapping(value = "/word35", method = RequestMethod.GET)
    public ModelAndView showWord35(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        String userName = request.getParameter("userName");
        //***************************卓正PageOffice组件的使用********************************
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dTitle = doc.openDataRegion("PO_title");
        //给数据区域赋值
        dTitle.setValue("某公司第二季度产量报表");
        //设置数据区域可编辑性
        dTitle.setEditing(false);//数据区域不可编辑

        DataRegion dA1 = doc.openDataRegion("PO_A_pro1");
        DataRegion dA2 = doc.openDataRegion("PO_A_pro2");
        DataRegion dB1 = doc.openDataRegion("PO_B_pro1");
        DataRegion dB2 = doc.openDataRegion("PO_B_pro2");

        //根据登录用户名设置数据区域可编辑性
        //A部门经理登录后
        if (userName.equals("zhangsan")) {
            userName = "A部门经理";
            dA1.setEditing(true);
            dA2.setEditing(true);
            dB1.setEditing(false);
            dB2.setEditing(false);
        }
        //B部门经理登录后
        else {
            userName = "B部门经理";
            dB1.setEditing(true);
            dB2.setEditing(true);
            dA1.setEditing(false);
            dA2.setEditing(false);
        }
        poCtrl.setWriter(doc);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        //设置保存页
        poCtrl.setSaveFilePage("/save/common");

        poCtrl.setMenubar(false);


        poCtrl.webOpen(dir + "test35.doc", OpenModeType.docSubmitForm, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word35");
        return mv;
    }


    @RequestMapping(value = "/word36", method = RequestMethod.GET)
    public ModelAndView showWord36(HttpServletRequest request, Map<String, Object> map) {


        ModelAndView mv = new ModelAndView("/word/Word36");
        return mv;
    }

    @RequestMapping(value = "/word37", method = RequestMethod.GET)
    public ModelAndView showWord37(HttpServletRequest request, Map<String, Object> map) {
        //--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        String userName = request.getParameter("userName");
        //***************************卓正PageOffice组件的使用********************************
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion d1 = doc.openDataRegion("PO_com1");
        DataRegion d2 = doc.openDataRegion("PO_com2");

        //给数据区域赋值
        d1.setValue("[word]word/content1.doc[/word]");
        d2.setValue("[word]word/content2.doc[/word]");

        //若要将数据区域内容存入文件中，则必须设置属性“setSubmitAsFile”值为true
        d1.setSubmitAsFile(true);
        d2.setSubmitAsFile(true);

        //根据登录用户名设置数据区域可编辑性
        //甲客户：zhangsan登录后
        if (userName.equals("zhangsan")) {
            d1.setEditing(true);
            d2.setEditing(false);
        }
        //乙客户：lisi登录后
        else {
            d2.setEditing(true);
            d1.setEditing(false);
        }

        poCtrl.setWriter(doc);

        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);

        //设置保存页
        poCtrl.setSaveDataPage("/save/doc/data19?userName=" + userName);


        poCtrl.webOpen(dir + "test37//test.doc", OpenModeType.docSubmitForm, userName);
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word37");
        return mv;
    }


}
