package com.example.demo.controller.word;

import com.zhuozhengsoft.pageoffice.*;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@Controller
public class WordController4 {
    @Value("${testPath}")
    private String dir;


   @RequestMapping(value="/word56", method= RequestMethod.GET)
    public ModelAndView showWord56(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
       poCtrl.addCustomToolButton("保存", "Save()", 1);
       poCtrl.addCustomToolButton("保存首页为图片", "SaveFirstAsImg()", 1);
       poCtrl.setSaveFilePage("/save/doc/data22");


        poCtrl.webOpen(dir+"test56.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word56");
        return mv;
    }
    @RequestMapping(value="/word57", method= RequestMethod.GET)
    public ModelAndView showWord57(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        Table table1 = doc.openDataRegion("PO_table").openTable(1);
        Cell  cell=table1.openCellRC(2,1);
        //删除坐标为(2,1)的单元格所在行
        table1.removeRowAt(cell);
        poCtrl.setCustomToolbar(false);
        poCtrl.setWriter(doc);


        poCtrl.webOpen(dir+"test57.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word57");
        return mv;
    }

    @RequestMapping(value="/word58", method= RequestMethod.GET)
    public ModelAndView showWord58(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        DataRegion mydr1 = doc.createDataRegion("PO_first", DataRegionInsertType.After, "[end]");
        mydr1.selectEnd();
        doc.insertPageBreak();//插入分页符
        DataRegion mydr2 = doc.createDataRegion("PO_second", DataRegionInsertType.After, "[end]");
        mydr2.setValue("[word]word/test2.doc[/word]");

        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.setWriter(doc);
//设置保存页面
        //poCtrl.setSaveFilePage("SaveFile.jsp");


        poCtrl.webOpen(dir+"test58.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word57");
        return mv;
    }

    @RequestMapping(value="/word59", method= RequestMethod.GET)
    public ModelAndView showWord59(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        Table table1 = doc.openDataRegion("PO_T001").openTable(1);

        table1.openCellRC(1,1).setValue("[image]images/logo.jpg[/image]");
        int dataRowCount = 5;//需要插入数据的行数
        int oldRowCount = 3;//表格中原有的行数
        // 扩充表格
        for (int j = 0; j < dataRowCount - oldRowCount; j++)
        {
            table1.insertRowAfter(table1.openCellRC(2, 5));  //在第2行的最后一个单元格下插入新行
        }
        // 填充数据
        int i = 1;
        while (i <= dataRowCount)
        {
            table1.openCellRC(i, 2).setValue("AA" + String.valueOf(i));
            table1.openCellRC(i, 3).setValue("BB" + String.valueOf(i));
            table1.openCellRC(i, 4).setValue("CC" + String.valueOf(i));
            table1.openCellRC(i, 5).setValue("DD" + String.valueOf(i));
            i++;
        }
        poCtrl.setWriter(doc);



        poCtrl.webOpen(dir+"test59.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word57");
        return mv;
    }

    @RequestMapping(value="/word60", method= RequestMethod.GET)
    public ModelAndView showWord60(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion = doc.openDataRegion("PO_regTable");
        //打开table，openTable(index)方法中的index代表Word文档中table位置的索引，从1开始
        Table table = dataRegion.openTable(1);

        //给table中的单元格赋值， openCellRC(int,int)中的参数分别代表第几行、第几列，从1开始
        table.openCellRC(3, 1).setValue("A公司");
        table.openCellRC(3, 2).setValue("开发部");
        table.openCellRC(3, 3).setValue("李清");

        //插入一行，insertRowAfter方法中的参数代表在哪个单元格下面插入一个空行
        table.insertRowAfter(table.openCellRC(3, 3));

        table.openCellRC(4, 1).setValue("B公司");
        table.openCellRC(4, 2).setValue("销售部");
        table.openCellRC(4, 3).setValue("张三");

        //设置表格行的高度
        table.setRowsHeight(30.5f);

        //设置表格的边框
        Border border=table.getBorder();
        // 设置边框的类型
        border.setBorderType(WdBorderType.wdFullGrid);//包含内边框
        //设置边框的颜色
        border.setLineColor(Color.red);
        //设置边框的线条样式
        border.setLineStyle(WdLineStyle.wdLineStyleDot );
        //设置边框的粗细
        border.setLineWidth(WdLineWidth.wdLineWidth150pt);

        //设置表格内字体样式
        com.zhuozhengsoft.pageoffice.wordwriter.Font font=dataRegion.getFont();
        //设置字体的是否加粗
        font.setBold(true);
        //设置字体的颜色
        font.setColor(Color.blue);
        //设置字体是否为斜体
        font.setItalic(true);
        //设置字体名称
        font.setName("宋体");
        //设置字体大小
        font.setSize(15.5f);

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);



        poCtrl.webOpen(dir+"test60.doc", OpenModeType.docNormalEdit,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word57");
        return mv;
    }

    @RequestMapping(value="/word61", method= RequestMethod.GET)
    public ModelAndView showWord61(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        poCtrl.addCustomToolButton("保存图片", "Save", 1);
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_image");
        dataRegion1.setEditing(true);//放图片的数据区域是可以编辑的，其它部分不可编辑
        poCtrl.setWriter(wordDoc);
        //设置保存页面
        poCtrl.setSaveDataPage("/save/doc/data23");



        poCtrl.webOpen(dir+"test61.doc", OpenModeType.docSubmitForm,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word61");
        return mv;
    }

    @RequestMapping(value="/word62", method= RequestMethod.GET)
    public ModelAndView showWord62(HttpServletRequest request, Map<String,Object> map){
//--- PageOffice的调用代码 开始 -----
        PageOfficeCtrl poCtrl=initPageOfficeCtrl(request);
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存","Save",1);
        poCtrl.addCustomToolButton("打印设置","PrintSet",0);
        poCtrl.addCustomToolButton("打印","PrintFile",6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        poCtrl.addCustomToolButton("-", "", 0);
        poCtrl.addCustomToolButton("关闭","Close",21);
//设置保存页面
        poCtrl.setSaveFilePage("/save/common");

//** 关键代码 禁止拷贝文档内容到外部 **
        poCtrl.setDisableCopyOnly(true);


        poCtrl.webOpen(dir+"test62.doc", OpenModeType.docSubmitForm,"zhangsan");
        map.put("pageoffice",poCtrl.getHtmlCode("PageOfficeCtrl1"));
        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("/word/Word62");
        return mv;
    }





    private  PageOfficeCtrl initPageOfficeCtrl(HttpServletRequest request){
        PageOfficeCtrl poCtrl=new PageOfficeCtrl(request);
        poCtrl.setServerPage("/poserver.zz");//设置授权程序servlet
        return poCtrl;
    }








}
