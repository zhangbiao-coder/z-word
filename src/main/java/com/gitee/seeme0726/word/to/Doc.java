package com.gitee.seeme0726.word.to;

import java.io.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import com.github.chengyuxing.common.io.IOutput;
import org.apache.commons.codec.binary.Base64;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.w3c.dom.Document;

public class Doc implements IOutput, AutoCloseable {
    private final InputStream inputStream;
    private final HWPFDocumentCore document;
    private final ByteArrayOutputStream outputStream;


    public Doc(InputStream inputStream) throws IOException {
        this.inputStream = inputStream;
        outputStream = new ByteArrayOutputStream();
        document = new HWPFDocument(inputStream);
    }

    /**
     * word2003
     */
    public Doc html() throws ParserConfigurationException, TransformerException, IOException {
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument()
        );
        //将图片转成base64的格式
        PicturesManager pictureRunMapper = (bytes, pictureType, s, v, v1) -> "data:image/png;base64," + Base64.encodeBase64String(bytes);
        wordToHtmlConverter.setPicturesManager(pictureRunMapper);
        //解析word文档
        wordToHtmlConverter.processDocument(document);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        DOMSource domSource = new DOMSource(htmlDocument);

        StreamResult streamResult = new StreamResult(outputStream);
        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        return this;
    }


    @Override
    public byte[] toBytes() {
        return outputStream.toByteArray();
    }

    @Override
    public void close() throws Exception {
        document.close();
        if (inputStream != null) {
            inputStream.close();
        }
        outputStream.close();

    }
}
