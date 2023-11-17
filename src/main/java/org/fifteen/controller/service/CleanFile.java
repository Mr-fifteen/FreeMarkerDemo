package org.fifteen.controller.service;

import java.io.File;

/**
 * @author liuwenxv
 * @date 2023/11/11 15:24
 * @comment
 */
public class CleanFile {
    public static void main(String[] args) {
        //文件路径
        String filepath = "D:\\shenzilong\\maven\\maven-repos";
        clean(new File(filepath) );

    }

    private static void clean(File file){
        File[] files = file.listFiles();
        if(files == null){
            return ;
        }
        for (int i = 0; i < files.length ; i++) {
            //如果是文件夹下钻
            if(files[i].isDirectory()){
                clean(files[i]);
            }
            String name = files[i].getName();
            //判断后缀不合规，则删除
            if(!(name.endsWith("pom")||name.endsWith("jar")||name.endsWith("sha1"))
                    && files[i].isFile() ){
                System.out.println("准备删除："+name);
                files[i].delete();
            }
        }
        return ;

    }

}
