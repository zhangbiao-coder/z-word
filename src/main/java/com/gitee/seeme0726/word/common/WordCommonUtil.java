package com.gitee.seeme0726.word.common;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

public  class WordCommonUtil {
    public static XWPFTableCell findCellByIndex(List<XWPFTable> tables,int[] tablePlaceholderIndex){
        int temTablePlaceholderIndex = tablePlaceholderIndex[0];
        int temRowPlaceholderIndex = tablePlaceholderIndex[1];
        int temCellPlaceholderIndex = tablePlaceholderIndex[2];
        XWPFTable findTable = tables.get(temTablePlaceholderIndex);
        XWPFTableRow findRow = findTable.getRow(temRowPlaceholderIndex);
        XWPFTableCell findCell = findRow.getCell(temCellPlaceholderIndex);
        return findCell;
    }
}
