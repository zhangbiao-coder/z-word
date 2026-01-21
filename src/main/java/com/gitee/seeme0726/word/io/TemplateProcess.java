package com.gitee.seeme0726.word.io;

import com.github.chengyuxing.common.io.IOutput;
import com.gitee.seeme0726.word.common.IoCommonUtil;
import com.gitee.seeme0726.word.common.TemplateUtil;
import fr.opensagres.xdocreport.core.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.*;

import static com.github.chengyuxing.common.utils.ObjectUtil.coalesce;
import static com.github.chengyuxing.common.utils.ObjectUtil.getDeepValue;

/**
 * 模板处理类
 */
public class TemplateProcess implements IOutput, AutoCloseable {

    private final XWPFDocument document;
    // 占位符参数map
    private final Map<String, Object> params = new HashMap<>();

    private final InputStream templateInputStream;

    /**
     * 构造函数
     *
     * @param inputStream 输入流
     * @throws IOException IOex
     */
    public TemplateProcess(InputStream inputStream) throws IOException {
        templateInputStream = inputStream;
        document = new XWPFDocument(templateInputStream);
    }

    /**
     * 增加文本占位符参数(一个)
     *
     * @param key   键
     * @param value 值
     * @return Builder对象
     */
    public TemplateProcess addParam(String key, Object value) {
        if (key != null && !key.isEmpty()) {
            params.put(key, value);
        }
        return this;
    }

    /**
     * 增加参数(多个)
     *
     * @param params 多个参数的map
     * @return Builder对象
     */
    public TemplateProcess addParams(Map<String, Object> params) {
        this.params.putAll(params);
        return this;
    }

    /**
     * 增加一个表格（适合模板行在表格中间的情况，就是表格处理有模板行，还有表头和表脚）
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置
     * @param placeholderLocation   指定模板行的下标位置
     * @param placeholderRows       指定模板行占几行
     * @param list                  集合数据
     * @return this
     */
    public TemplateProcess addTable(int tablePlaceholderIndex, int placeholderLocation, int placeholderRows, Iterable<Map<String, Object>> list) {
        List<XWPFTable> tables = document.getTables();

        //找到指定表格
        XWPFTable table = tables.get(tablePlaceholderIndex);

        TemplateUtil.replaceTemplateRows(table, placeholderLocation, placeholderRows, list);
        return this;
    }

    /**
     * 根据键名从参数对象中获取一个集合数据增加一个表格
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置
     * @param placeholderLocation   指定模板行的下标位置
     * @param placeholderRows       指定模板行占几行
     * @param key                   来自于参数的键名
     * @return this
     */
    @SuppressWarnings("unchecked")
    public TemplateProcess setTable(int tablePlaceholderIndex, int placeholderLocation, int placeholderRows, String key) {
        List<XWPFTable> tables = document.getTables();

        //找到指定表格
        XWPFTable table = tables.get(tablePlaceholderIndex);

        Object list = coalesce(getDeepValue(params, key), Collections.emptyList());
        if (!(list instanceof Iterable<?>)) {
            throw new IllegalArgumentException(key + " is not a Iterable");
        }

        TemplateUtil.replaceTemplateRows(table, placeholderLocation, placeholderRows, (Iterable<Map<String, Object>>) list);
        return this;
    }

    /**
     * 增加一个表格（适合模板行在表格中间的情况，就是表格处理有模板行，还有表头和表脚）
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置 数组（第一个只是文档中的第几个表格，第二个值是表格中的第几行，第三个值表示第几个单元格）
     * @param placeholderLocation   指定模板行的下标位置
     * @param placeholderRows       指定模板行占几行
     * @param list                  集合数据
     * @return this
     */
    public TemplateProcess addTable(int[] tablePlaceholderIndex, int placeholderLocation, int placeholderRows, Iterable<Map<String, Object>> list) {
        List<XWPFTable> tables = document.getTables();

        //找到指定表格
        XWPFTableCell findCell = TemplateUtil.findCellByIndex(tables, tablePlaceholderIndex);
        XWPFTable table = findCell.getTables().get(0);

        TemplateUtil.replaceTemplateRows(table, placeholderLocation, placeholderRows, list);
        return this;
    }


    /**
     * 增加图片（指定图片在文档中的宽高）
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置 数组（第一个只是文档中的第几个表格，第二个值是表格中的第几行，第三个值表示第几个单元格）
     * @param inputStream           图片字节流
     * @param width                 宽度
     * @param height                高度
     * @param fileName              文件名（必须以图片格式结尾）
     * @return this
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public TemplateProcess addImg(int[] tablePlaceholderIndex, InputStream inputStream, int width, int height, String fileName) throws IOException, InvalidFormatException {
        List<XWPFTable> tables = document.getTables();
        //找到指定表格
        XWPFTableCell findCell = TemplateUtil.findCellByIndex(tables, tablePlaceholderIndex);
        TemplateUtil.insertImg(findCell, inputStream, width, height, fileName);
        return this;
    }

    /**
     * 增加图片（动态的调整图片在文档中的宽高，指定图片的格式）
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置 数组（第一个只是文档中的第几个表格，第二个值是表格中的第几行，第三个值表示第几个单元格）
     * @param inputStream           图片字节流
     * @param fileName              文件名（必须以图片格式结尾）
     * @return this
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public TemplateProcess addImg(int[] tablePlaceholderIndex, InputStream inputStream, String fileName) throws IOException, InvalidFormatException {
        //获取图片文件流
        //计算适合文档宽高的图片EMU数值
        BufferedImage read = null;
        //由于inputstream 流不能进行二次操作
        try (ByteArrayOutputStream baos = IoCommonUtil.cloneInputStream(inputStream);
             InputStream stream1 = new ByteArrayInputStream(baos.toByteArray());
             InputStream stream2 = new ByteArrayInputStream(baos.toByteArray());) {


            read = ImageIO.read(stream1);
            int width = Units.toEMU(read.getWidth());
            int height = Units.toEMU(read.getHeight());

            //1 EMU = 1/914400英寸= 1/36000 mm,15是word文档中图片能设置的最大宽度cm
            if (width / 360000 > 15) {
                NumberFormat f = NumberFormat.getNumberInstance();
                f.setMaximumFractionDigits(0);
                f.setRoundingMode(RoundingMode.UP);
                double d = width / 360000d / 15d;
                width = Integer.parseInt(f.format(width / d).replace(",", ""));
                height = Integer.parseInt(f.format(height / d).replace(",", ""));
            }

            addImg(tablePlaceholderIndex, stream2, width, height, fileName);
            return this;
        } finally {
            if (read != null) {
                read.flush();
            }
        }

    }


    /**
     * 增加图片（动态的调整图片在文档中的宽高，默认图片为png格式）
     *
     * @param tablePlaceholderIndex 指定模板表格的下标位置 数组（第一个只是文档中的第几个表格，第二个值是表格中的第几行，第三个值表示第几个单元格）
     * @param inputStream           图片字节流
     * @return this
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public TemplateProcess addImg(int[] tablePlaceholderIndex, InputStream inputStream) throws IOException, InvalidFormatException {
        addImg(tablePlaceholderIndex, inputStream, "default.png");
        return this;
    }

    /**
     * 替换文档中的指定下标图片
     *
     * @param pictureIndex   要替换的图片的下标序号，从0开始
     * @param newImageStream 新图片的输入流
     * @return this
     * @throws IOException            IOException
     * @throws InvalidFormatException InvalidFormatException
     */
    public TemplateProcess replaceImg(int pictureIndex, InputStream newImageStream) throws IOException, InvalidFormatException {
        // 1. 收集文档中所有图片及其相关信息
        List<XWPFPictureData> allPictures = document.getAllPictures();
        // 2. 检查索引有效性
        if (pictureIndex < 0 || pictureIndex >= allPictures.size()) {
            throw new IndexOutOfBoundsException("图片索引超出范围。文档包含 " + allPictures.size() + " 张图片。");
        }
        // 3. 获取目标图片信息
        XWPFPictureData pictureData = allPictures.get(pictureIndex);
        PackagePart packagePart = pictureData.getPackagePart();
        // 4. 用新图片覆盖原始数据 (保持相同格式!)
        try (OutputStream outputStream = packagePart.getOutputStream()) {
            IOUtils.copy(newImageStream, outputStream);
        }
        return this;
    }


    @Override
    public byte[] toBytes() throws IOException {
        try (ByteArrayOutputStream stream = new ByteArrayOutputStream()) {
            TemplateUtil.changeText(document, params);
            document.write(stream);
            return stream.toByteArray();
        }
    }

    @Override
    public void close() throws Exception {
        document.close();
        templateInputStream.close();
    }
}
