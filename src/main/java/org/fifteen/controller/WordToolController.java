package org.fifteen.controller;


import fr.opensagres.xdocreport.core.XDocReportException;
import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.fifteen.controller.bean.Goods;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;

import java.util.ArrayList;
import java.util.List;

/**
 * @author liuwenxv
 * @date 2023/1/15 19:56
 * @comment
 */

@RequestMapping
@RestController
public class WordToolController {

    @RequestMapping("/test")
    public String test() throws IOException, XDocReportException {
        generateWord();
        return "Hello World!";
    }


    public void generateWord() throws IOException, XDocReportException {
        //获取Word模板，模板存放路径在项目的resources目录下
        InputStream ins = this.getClass().getResourceAsStream("/模板.docx");
//        InputStream ins = this.getClass().getResourceAsStream("/模板.ftl");
        //注册xdocreport实例并加载FreeMarker模板引擎
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(ins, TemplateEngineKind.Freemarker);


        //创建xdocreport上下文对象
        IContext context = report.createContext();

        //创建要替换的文本变量
        context.put("city", "北京市");
        context.put("startDate", "2020-09-17");
        context.put("endDate", "2020-10-16");
        context.put("totCnt", 3638763);
        context.put("totAmt", "6521");
        context.put("onCnt", 2874036);
        context.put("onAmt", "4768");
        context.put("offCnt", 764727);
        context.put("offAmt", "1753");
        context.put("typeCnt", 36);

        List<Goods> goodsList = new ArrayList<Goods>();
        Goods goods1 = new Goods();
        goods1.setNum(1);
        goods1.setType("臭美毁肤");
        goods1.setSv(675512);
        goods1.setSa("589");
        Goods goods2 = new Goods();
        goods2.setNum(2);
        goods2.setType("女装");
        goods2.setSv(602145);
        goods2.setSa("651");
        Goods goods3 = new Goods();
        goods3.setNum(3);
        goods3.setType("手机");
        goods3.setSv(587737);
        goods3.setSa("866");
        Goods goods4 = new Goods();
        goods4.setNum(4);
        goods4.setType("家具家装");
        goods4.setSv(551193);
        goods4.setSa("783");
        Goods goods5 = new Goods();
        goods5.setNum(5);
        goods5.setType("食物饮品");
        goods5.setSv(528604);
        goods5.setSa("405");
        goodsList.add(goods1);
        goodsList.add(goods2);
        goodsList.add(goods3);
        goodsList.add(goods4);
        goodsList.add(goods5);
        context.put("goods", goodsList);

        //创建字段元数据
        FieldsMetadata fm = report.createFieldsMetadata();
        //Word模板中的表格数据对应的集合类型
        fm.load("goods", Goods.class, true);

        //输出到本地目录
        FileOutputStream out = new FileOutputStream(new File("D://商品销售报表.docx"));
        report.process(context, out);

        ins.close();
        out.close();
    }


    public static void main(String[] args) throws IOException {

        //原word文件
//        String sourcePath = "resource\word2pdf\数据申请表.docx";
        String sourcePath = "D:\\fileDemo\\数据申请表.docx";
        //导出文件名
        String targetPath = "D:\\fileDemo\\数据申请表2.pdf";

        docx4jWordToPdf( sourcePath,  targetPath);

//        FileInputStream fileInputStream = new FileInputStream("D:\\助学贷款申请表.docx");
//        XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
//        PdfOptions pdfOptions = PdfOptions.create();
//        FileOutputStream fileOutputStream = new FileOutputStream("D:\\助学贷款申请表.pdf");
//        PdfConverter.getInstance().convert(xwpfDocument,fileOutputStream,pdfOptions);
//        fileInputStream.close();
//        fileOutputStream.close();

    }

    /**
     * 通过docx4j 实现word转pdf
     *
     * @param sourcePath 源文件地址 如 /root/example.doc
     * @param targetPath 目标文件地址 如 /root/example.pdf
     */
    public static void docx4jWordToPdf(String sourcePath, String targetPath) {
        try {
            //doc转xml
            File fileXML=null;
            if(sourcePath != null && sourcePath.endsWith("docx") ){
                File fileDocx =   new File(sourcePath);
//                String[] file = sourcePath.split(".");
                String[] file = sourcePath.split("\\.");
                fileXML =   new File(file[0]+".xml");
                FileUtils.copyFile(fileDocx,fileXML);
            }else{
                fileXML= new File(sourcePath);
            }


            WordprocessingMLPackage pkg = Docx4J.load(fileXML);
            Mapper fontMapper = new IdentityPlusMapper();
            fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
            fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
            fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
            fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
            fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
            fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
            fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
            fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
            fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));

            fontMapper.put("方正小标宋简体", PhysicalFonts.get("FangSong"));

            fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
            fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
            fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));
            fontMapper.put("等线", PhysicalFonts.get("SimSun"));
            fontMapper.put("等线 Light", PhysicalFonts.get("SimSun"));
            fontMapper.put("华文琥珀", PhysicalFonts.get("STHupo"));
            fontMapper.put("华文隶书", PhysicalFonts.get("STLiti"));
            fontMapper.put("华文新魏", PhysicalFonts.get("STXinwei"));
            fontMapper.put("华文彩云", PhysicalFonts.get("STCaiyun"));
            fontMapper.put("方正姚体", PhysicalFonts.get("FZYaoti"));
            fontMapper.put("方正舒体", PhysicalFonts.get("FZShuTi"));
            fontMapper.put("华文细黑", PhysicalFonts.get("STXihei"));
            fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
            fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
            fontMapper.put("新細明體", PhysicalFonts.get("SimSun"));
            pkg.setFontMapper(fontMapper);

            Docx4J.toPDF(pkg, new FileOutputStream(targetPath));
        } catch (Exception e) {
            System.out.println("[docx4j] word转pdf失败:{}"+e.toString());
//            LOGGER.error();
        }
    }





}
