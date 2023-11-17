package org.fifteen.controller.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

//import com.itextpdf.text.Document;
//import com.itextpdf.text.DocumentException;
//import com.itextpdf.text.pdf.PdfWriter;
//import com.itextpdf.text.pdf.PdfReader;
//import org.apache.poi.xwpf.converter.pdf.PdfConverter;
//import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import com.aspose.words.PdfSaveOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
/**
 * @author liuwenxv
 * @date 2023/2/17 20:46
 * @comment
 */
public class WordToPDF2 {

    /**
     * 待确定，不知道当初有没有运行成功过
     * @param args
     * @throws IOException
     */
    public static void main(String[] args) throws IOException{

        String source = "C:\\Users\\Test\\Desktop\\sample.docx";
        String dest = "Output.pdf";

        // 由Word创建XWPFDocument对象
        XWPFDocument document = new XWPFDocument(new FileInputStream(source));

        // 准备好一个pdf的输出流
        FileOutputStream out = new FileOutputStream(dest);

        // 将XWPFDocument转换成PDF
//        PdfOptions pdfOptions = PdfOptions.create();
//        PdfConverter.getInstance().convert(document, out, pdfOptions);

        System.out.println("Word文档转换成功！");
    }

}
