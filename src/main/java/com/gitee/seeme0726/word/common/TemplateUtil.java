package com.gitee.seeme0726.word.common;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 处理模板占位符的工具类
 */
public class TemplateUtil {
    private final static String pattern = "\\$\\{[^}]+}";
    private static final Pattern compile = Pattern.compile(pattern);

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, String> textMap) {
        // 获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        replaceParagraphsText(paragraphs, textMap);

        //表格中的站位符也要处理
        //得到word中的表格
        Iterator<XWPFTable> it = document.getTablesIterator();
        while (it.hasNext()) {
            XWPFTable table = it.next();
            List<XWPFTableRow> rows = table.getRows();
            //读取每一行数据
            for (XWPFTableRow row : rows) {
                //读取每一列数据
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    List<XWPFParagraph> cellParagraphs = cell.getParagraphs();
                    replaceParagraphsText(cellParagraphs, textMap);
                }
            }
        }
    }

    /**
     * 替换段落占位符的值
     *
     * @param paragraphs 段落
     * @param textMap    数据集
     */
    public static void replaceParagraphsText(List<XWPFParagraph> paragraphs, Map<String, String> textMap) {
        for (XWPFParagraph paragraph : paragraphs) {
            // 获取到段落中的所有文本内容
            String text = paragraph.getText();
            // 判断此段落中是否有需要进行替换的文本
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //原值
                    String rtext = run.text();
                    // 替换模板原来位置
                    Matcher m = compile.matcher(rtext);
                    if (m.find()) {
                        String key = m.group();
                        String value = textMap.get(getKey(key));
                        value = value == null ? "--" : value;
                        if (value.indexOf("\n") > 0) {
                            String[] texts = value.split("\n");
                            for (int i = 0; i < texts.length; i++) {
                                String s = texts[i];
                                if (i == 0) {
                                    run.setText(rtext.replace(key, s), 0);
                                } else {
                                    run.addBreak();
                                    run.setText(s, i);
                                }
                            }
                        } else {
                            run.setText(rtext.replace(key, value), 0);
                        }
                    }
                }
            }
        }
    }

    /**
     * 替换表格占位符号并复制行
     *
     * @param table               包含模板行的表格
     * @param placeholderLocation 指定模板行的下标位置
     * @param placeholderRows     指定模板行占几行
     * @param list                数据集
     */
    public static void replaceTemplateRows(XWPFTable table, int placeholderLocation, int placeholderRows, List<Map<String, String>> list) {
        //获取模板行
        List<XWPFTableRow> temRows = new ArrayList<>();
        for (int i = 0; i < placeholderRows; i++) {
            temRows.add(table.getRow(placeholderLocation + i));
        }

        for (int i = 0; i < list.size(); i++) {
            //当前行数据集
            Map<String, String> map = list.get(i);

            //复制模板行
            for (XWPFTableRow temRow : temRows) {
                //复制模板行得到新行
                XmlObject copy = temRow.getCtRow().copy();
                XWPFTableRow newRow = new XWPFTableRow((CTRow) copy, table);
                //替换新行中的展位符
                List<XWPFTableCell> cells = newRow.getTableCells();
                for (XWPFTableCell cell : cells) {
                    List<XWPFParagraph> cellParagraphs = cell.getParagraphs();
                    TemplateUtil.replaceParagraphsText(cellParagraphs, map);
                }
                //把新行插入表格，插入下标为（模板行下标 + 模板行数 + 当前遍历下标）
                table.addRow(newRow, placeholderLocation + placeholderRows + i);

            }
        }

        //最后移除模板行
        for (int i = 0; i < placeholderRows; i++) {
            table.removeRow(placeholderLocation + i);
        }
    }


    /**
     * 在指定单元格插入图片 （手动指定图片宽高）
     *
     * @param cell        指定单元格
     * @param inputStream 图片输入流
     * @param width       宽度
     * @param height      高度
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public static XWPFPicture insertImg(XWPFTableCell cell, InputStream inputStream, int width, int height, String fileName) throws IOException, InvalidFormatException {
        try {
            //判断图片的格式
            int pictureType ;
            if (fileName.endsWith(".emf")) {
                pictureType = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (fileName.endsWith(".wmf")) {
                pictureType = XWPFDocument.PICTURE_TYPE_WMF;
            } else if (fileName.endsWith(".pict")) {
                pictureType = XWPFDocument.PICTURE_TYPE_PICT;
            } else if (fileName.endsWith(".jpeg") || fileName.endsWith(".jpg")) {
                pictureType = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (fileName.endsWith(".png")) {
                pictureType = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (fileName.endsWith(".dib")) {
                pictureType = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (fileName.endsWith(".gif")) {
                pictureType = XWPFDocument.PICTURE_TYPE_GIF;
            } else if (fileName.endsWith(".tiff")) {
                pictureType = XWPFDocument.PICTURE_TYPE_TIFF;
            } else if (fileName.endsWith(".eps")) {
                pictureType = XWPFDocument.PICTURE_TYPE_EPS;
            } else if (fileName.endsWith(".bmp")) {
                pictureType = XWPFDocument.PICTURE_TYPE_BMP;
            } else if (fileName.endsWith(".wpg")) {
                pictureType = XWPFDocument.PICTURE_TYPE_WPG;
            } else {
                throw new RuntimeException("Unsupported picture: " + fileName +
                        ". Expected emf|wmf|pict|jpeg|png|dib|gif|tiff|eps|bmp|wpg");
            }

            //获取单元格的段落
            XWPFParagraph paragraphs = cell.getParagraphs().get(0).isEmpty() ? cell.getXWPFDocument().createParagraph() : cell.getParagraphs().get(0);
            XWPFRun run = paragraphs.getRuns().isEmpty() ? paragraphs.createRun() : paragraphs.getRuns().get(0);
            return run.addPicture(inputStream, pictureType, fileName, width, height);

        } finally {
            if (inputStream != null) {
                inputStream.close();
            }

        }

    }


    public static String getKey(String g) {
        return g.substring(2, g.length() - 1);
    }

    /**
     * 判断文本中是否包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        return text.contains("$");
    }

}
