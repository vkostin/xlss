package io.github.vkostin.xlss.sheet;

import io.github.vkostin.xlss.styles.DefaultColors;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.function.Function;

import static io.github.vkostin.xlss.styles.StylesHelper.setBgColor;
import static io.github.vkostin.xlss.styles.StylesHelper.setNoBorder;

public class SheetHelpers {
    public void setDefaultSheetStyle(XSSFWorkbook workbook, Sheet sheet, Function<DefaultStyleRendererParams, XSSFCellStyle> createDefaultStyle) {
        if (createDefaultStyle != null) {
            XSSFCellStyle defaultStyle = createDefaultStyle.apply(new DefaultStyleRendererParams(workbook, sheet));
            if (defaultStyle != null) {
                org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol cTCol =
                        ((XSSFSheet) sheet).getCTWorksheet().getColsArray(0).addNewCol();
                cTCol.setMin(1);
                cTCol.setMax(16384);
                cTCol.setWidth(12.7109375);
                cTCol.setStyle(defaultStyle.getIndex());
            }
        }
    }

    public record DefaultStyleRendererParams(XSSFWorkbook workbook, Sheet sheet) {

    }

    public static Function<DefaultStyleRendererParams, XSSFCellStyle> getWhiteBgStyle() {
        return (DefaultStyleRendererParams params) -> {
            XSSFCellStyle defaultStyle = params.workbook.createCellStyle();
            DefaultIndexedColorMap indexedColorMap = new DefaultIndexedColorMap();
            defaultStyle.setAlignment(HorizontalAlignment.LEFT);
            defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setNoBorder(defaultStyle);
            setBgColor(defaultStyle, DefaultColors.white(), indexedColorMap);
            return defaultStyle;
        };
    }
}
