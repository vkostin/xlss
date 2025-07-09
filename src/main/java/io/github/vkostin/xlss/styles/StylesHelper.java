package io.github.vkostin.xlss.styles;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;

public class StylesHelper {
    public static void setThinBorder(XSSFCellStyle style) {
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
    }

    public static void setNoBorder(XSSFCellStyle style) {
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderTop(BorderStyle.NONE);
    }

    public static void setBgColor(XSSFCellStyle style, Color color, DefaultIndexedColorMap indexedColorMap) {
        style.setFillForegroundColor(new XSSFColor(color, indexedColorMap));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }
}
