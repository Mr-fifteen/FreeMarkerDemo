package org.fifteen.controller.service;

/**
 * @author liuwenxv
 * @date 2023/2/16 14:24
 * @comment
 */

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WordToPDF
{
    public static void main(String [] args) throws Exception {
        toPDF();
    }


//    <CtClass> void crackAsposeWords() throws Exception {
//        ClassPool.getDefault().insertClassPath("C:\\Users\\coder\\Downloads\\aspose-words-20.12-jdk17.jar");
//        CtClass zzZJJClass = ClassPool.getDefault().getCtClass("com.aspose.words.zzZDZ");
//        CtMethod zzZ4u = zzZJJClass.getDeclaredMethod("zzZ4n");
//        CtMethod zzZ4t = zzZJJClass.getDeclaredMethod("zzZ4m");
//        zzZ4u.setBody("{return 1;}");
//        zzZ4t.setBody("{return 1;}");
//        zzZJJClass.writeFile("C:\\Users\\coder\\Desktop\\");
//    }


    //转存后的文件有水印
    static void  toPDF() throws Exception{
        // Load the Word Document
        Document doc = new Document(new FileInputStream("D:\\fileDemo\\数据申请表.docx"));

        // Save the document in PDF format
        doc.save(new FileOutputStream("D:\\fileDemo\\数据申请表.pdf"), SaveFormat.PDF);
    }

}
