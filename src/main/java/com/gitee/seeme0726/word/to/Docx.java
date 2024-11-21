package com.gitee.seeme0726.word.to;

import com.github.chengyuxing.common.io.IOutput;
import fr.opensagres.poi.xwpf.converter.xhtml.Base64EmbedImgManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class Docx implements IOutput, AutoCloseable {
    private final InputStream inputStream;
    private final XWPFDocument document;
    private final ByteArrayOutputStream outputStream;

    public Docx(InputStream inputStream) throws IOException {
        this.inputStream = inputStream;
        outputStream = new ByteArrayOutputStream();
        document = new XWPFDocument(inputStream);
    }

    /**
     * word2003
     */
    public Docx html() throws  IOException {
        XHTMLOptions xhtmlOptions = XHTMLOptions.create();
        // 是否忽略未使用的样式
        xhtmlOptions.setIgnoreStylesIfUnused(false);
        // 设置片段模式，<div>标签包裹
        xhtmlOptions.setFragment(true);
        // 图片转base64
        xhtmlOptions.setImageManager(new Base64EmbedImgManager());
        // 转换htm1
        XHTMLConverter.getInstance().convert(document, outputStream, xhtmlOptions);
        return this;
    }


    @Override
    public byte[] toBytes() {
        return outputStream.toByteArray();
    }

    @Override
    public void close() throws Exception {
        document.close();
        inputStream.close();
        outputStream.close();

    }
}
