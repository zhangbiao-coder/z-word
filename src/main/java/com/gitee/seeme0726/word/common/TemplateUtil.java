package com.gitee.seeme0726.word.common;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.github.chengyuxing.common.utils.ObjectUtil.coalesce;
import static com.github.chengyuxing.common.utils.ObjectUtil.getDeepValue;

/**
 * 处理模板占位符的工具类
 */
public class TemplateUtil {
    @SuppressWarnings("UnnecessaryUnicodeEscape")
    private static final char BREAK_RUN_CHAR = '\u0e3a';
    /**
     * 匹配 $|{ | use|r.id | }
     */
    private static final Pattern STRING_TEMPLATE_HOLDER_PATTERN = Pattern.compile("\\$" + BREAK_RUN_CHAR + "?\\{[\\s" + BREAK_RUN_CHAR + "]*(?<key>[\\w." + BREAK_RUN_CHAR + "]+)[\\s" + BREAK_RUN_CHAR + "]*}");

    /**
     * 替换段落文本
     *
     * @param document docx解析对象
     * @param textMap  需要替换的信息集合
     */
    public static void changeText(XWPFDocument document, Map<String, Object> textMap) {
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
     * @param args       数据集
     */
    public static void replaceParagraphsText(List<XWPFParagraph> paragraphs, Map<String, Object> args) {
        for (XWPFParagraph paragraph : paragraphs) {
            // 获取到段落中的所有文本内容
            String content = paragraph.getText();
            // 判断此段落中是否有需要进行替换的文本
            if (checkText(content)) {
                List<XWPFRun> runs = paragraph.getRuns();
                String[] newRuns = formatRuns(runs, args);
                if (newRuns.length <= runs.size()) {
                    for (int i = 0; i < newRuns.length; i++) {
                        String nr = newRuns[i];
                        XWPFRun run = runs.get(i);
                        if (nr.indexOf("\n") > 0) {
                            String[] texts = nr.split("\n");
                            for (int j = 0; j < texts.length; j++) {
                                String s = texts[j];
                                if (j == 0) {
                                    run.setText(s, 0);
                                } else {
                                    run.addBreak();
                                    run.setText(s, j);
                                }
                            }
                        } else {
                            run.setText(nr, 0);
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
    public static void replaceTemplateRows(XWPFTable table, int placeholderLocation, int placeholderRows, Iterable<Map<String, Object>> list) {
        //获取模板行
        List<XWPFTableRow> temRows = new ArrayList<>();
        for (int i = 0; i < placeholderRows; i++) {
            temRows.add(table.getRow(placeholderLocation + i));
        }

        int i = 0;
        for (Map<String, Object> map : list) {
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
                i++;
            }
        }

        //最后移除模板行
        for (int j = 0; j < placeholderRows; j++) {
            table.removeRow(placeholderLocation + j);
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
            int pictureType;
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

    /**
     * 判断文本中是否包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        return text.contains("$");
    }

    public static XWPFTableCell findCellByIndex(List<XWPFTable> tables, int[] tablePlaceholderIndex) {
        int temTablePlaceholderIndex = tablePlaceholderIndex[0];
        int temRowPlaceholderIndex = tablePlaceholderIndex[1];
        int temCellPlaceholderIndex = tablePlaceholderIndex[2];
        XWPFTable findTable = tables.get(temTablePlaceholderIndex);
        XWPFTableRow findRow = findTable.getRow(temRowPlaceholderIndex);
        XWPFTableCell findCell = findRow.getCell(temCellPlaceholderIndex);
        return findCell;
    }

    private static String[] formatRuns(List<XWPFRun> runs, Map<String, Object> args) {
        String brc = String.valueOf(BREAK_RUN_CHAR);
        StringJoiner sb = new StringJoiner(brc);
        for (XWPFRun r : runs) {
            sb.add(r.text());
        }
        Matcher m = STRING_TEMPLATE_HOLDER_PATTERN.matcher(sb.toString());
        StringBuffer newSb = new StringBuffer();
        while (m.find()) {
            int cc = timesOfSubstring(m.group(), brc);
            String key = m.group("key").replace(brc, "");
            StringBuilder value = new StringBuilder(coalesce(getDeepValue(args, key), "--").toString());
            while (cc > 0) {
                value.append(BREAK_RUN_CHAR);
                cc--;
            }
            m.appendReplacement(newSb, Matcher.quoteReplacement(value.toString()));
        }
        m.appendTail(newSb);
        return newSb.toString().split(brc, -1);
    }

    private static int timesOfSubstring(String str, String sub) {
        if (sub.isEmpty()) {
            return 0;
        }
        int times = 0;
        if (sub.length() == 1) {
            char c = sub.charAt(0);
            for (int i = 0; i < str.length(); i++) {
                if (str.charAt(i) == c) {
                    times++;
                }
            }
            return times;
        }

        int fromIndex = 0;
        while ((fromIndex = str.indexOf(sub, fromIndex)) != -1) {
            times++;
            fromIndex += sub.length();
        }
        return times;
    }
}
