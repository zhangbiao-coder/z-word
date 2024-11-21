package com.gitee.seeme0726.word;

import com.gitee.seeme0726.word.io.TemplateProcess;
import com.gitee.seeme0726.word.to.Doc;
import com.gitee.seeme0726.word.to.Docx;
import com.github.chengyuxing.common.io.FileResource;

import java.io.IOException;
import java.io.InputStream;

/**
 * 操作word入口类
 */
public final class Words {


    /**
     * modules: 根据模板导出word
     * 传入入模板路径
     * 如果传入resource目录下的文件路径
     */
    public static TemplateProcess ofTemplate(String path) throws IOException {
        return ofTemplate(new FileResource(path).getInputStream());
    }

    /**
     * modules: 根据模板导出word
     * 传入模板路径
     * 如果传入的直接是文件輸入流
     */
    public static TemplateProcess ofTemplate(InputStream inputStream) throws IOException {
        return new TemplateProcess(inputStream);
    }

    /**
     * modules: doc转换为各种格式
     * 传入的是文本路径
     */
    public static Doc docTo(String path) throws IOException {
        return new Doc(new FileResource(path).getInputStream());
    }

    /**
     * modules: doc转换为各种格式
     * 传入的是word文件輸入流
     */
    public static Doc docTo(InputStream inputStream) throws IOException {
        return new Doc(inputStream);
    }

    /**
     * modules: docx转换为各种格式
     * 传入的是文本路径
     */
    public static Docx docxTo(String path) throws IOException {
        return new Docx(new FileResource(path).getInputStream());
    }

    /**
     * modules: docx转换为各种格式
     * 传入的是word文件輸入流
     */
    public static Docx docxTo(InputStream inputStream) throws IOException {
        return new Docx(inputStream);
    }

}
