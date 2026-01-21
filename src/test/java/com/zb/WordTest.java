package com.zb;

import com.github.chengyuxing.common.DataRow;
import com.github.chengyuxing.common.io.FileResource;
import com.gitee.seeme0726.word.Words;
import com.gitee.seeme0726.word.io.TemplateProcess;
import com.gitee.seeme0726.word.to.Doc;
import com.gitee.seeme0726.word.to.Docx;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.util.*;
import java.util.stream.IntStream;

public class WordTest {

    @Test
    public void docToHtml() throws Exception {
        try (Doc doc = Words.docTo(new FileResource("file:/D:/ces.doc").getInputStream())) {
            doc.html().saveTo("D:/ces.html");
        }
    }

    @Test
    public void docxToHtml() throws Exception {
        try (Docx docx = Words.docxTo(new FileResource("file:/C:/temp/zwordTest.docx").getInputStream())) {
            docx.html().saveTo("C:/temp/zwordTest.html");
        }
    }

    @Test
    public void xWord2() throws Exception {
        List<Map<String, Object>> list = getRes();
        try (TemplateProcess templateProcess = Words.ofTemplate(new FileResource("file:/C:/temp/zwordTest.docx").getInputStream())) {
            templateProcess
                    .addParam("a", "測試到處\nl\n我在测试啦啦啦")
                    .addParam("b", "bbbbbbbbbbbbbbbbbbbbbbb")
                    .addParam("c", "ccccccccccccccccccccccc")
                    .addImg(new int[]{0, 0, 0}, new FileResource("file:/C:/Users/张彪/Pictures/1087462.jpg").getInputStream())
                    .saveTo("C:/temp/zwordTestSuccess.docx");
        }
    }

    @Test
    public void yWord2() throws Exception {
        DataRow row = DataRow.of(
                "user", DataRow.of("name", "cyx", "address", Arrays.asList("a", "b", "c")),
                "password", "12345\n67890",
                "database", "oracle");
        try (TemplateProcess templateProcess = Words.ofTemplate(new FileResource("file:/Users/chengyuxing/Downloads/mysql配置.docx").getInputStream())) {
            templateProcess
                    .addParams(row)
                    .setTable(0, 0, 0, "user")
                    .saveTo("/Users/chengyuxing/Downloads/mysql配置结果.docx");
        }
    }

    @Test
    public void xWord() throws Exception {
        List<Map<String, Object>> list = getRes();
        try (TemplateProcess templateProcess = Words.ofTemplate(new FileResource("file:/D:/tem.docx").getInputStream())) {
            templateProcess.addParam("b", "測試到處\nl\n我在测试啦啦啦")
                    .addTable(new int[]{0, 2, 0}, 1, 1, list)
                    .addImg(new int[]{0, 0, 0}, new FileResource("file:/D:/1.jpg").getInputStream(), "1.jpg")
                    .saveTo("D:/b.docx");
        }
    }

    private List<Map<String, Object>> getRes() {
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("xh", "1");
        map.put("xm", "张三\n涨到");
        map.put("xb", "男");
        map.put("bz", "张三是个英雄");
        list.add(map);
        Map<String, Object> map1 = new HashMap<>();
        map1.put("xh", "2");
        map1.put("xm", "李四");
        map1.put("xb", "女");
        map1.put("bz", "李四是个张三媳妇儿");
        list.add(map1);
        Map<String, Object> map2 = new HashMap<>();
        map2.put("xh", "3");
        map2.put("xm", "王五");
        map2.put("xb", "男");
        map2.put("bz", "王五是个大坏蛋");
        list.add(map2);
        return list;
    }


    @Test
    public void xWord3() throws Exception {
        List<Map<String, Object>> list = getRes();
        try (TemplateProcess templateProcess = Words.ofTemplate(new FileResource("file:/C:/temp/zwordTest.docx").getInputStream())) {
            templateProcess
                    .addParam("a", "測試到處\nl\n我在测试啦啦啦")
                    .addParam("b", "bbbbbbbbbbbbbbbbbbbbbbb")
                    .addParam("c", "ccccccccccccccccccccccc")
                    .addImg(new int[]{0, 0, 0}, new FileResource("file:/C:/temp/2.png").getInputStream())
                    .replaceImg(0, new FileResource("file:/C:/temp/1.gif").getInputStream())
                    .saveTo("C:/temp/zwordTestSuccess.docx");
        }
    }

    @Test
    public void xWord4() throws Exception {
        List<Map<String, Object>> list = getRes();
        try (TemplateProcess templateProcess = Words.ofTemplate(new FileResource("file:/C:/temp/zwordTest.docx").getInputStream())) {
            templateProcess
                    .addParam("a", "測試到處\nl\n我在测试啦啦啦")
                    .addParam("b", "bbbbbbbbbbbbbbbbbbbbbbb")
                    .addParam("c", "ccccccccccccccccccccccc")
                    .addImg(new int[]{0, 0, 0}, new FileResource("file:/C:/temp/2.png").getInputStream())
                    .replaceImg(0, new ByteArrayInputStream(new byte[0]))
                    .saveTo("C:/temp/zwordTestSuccess.docx");
        }
    }


}
