package io.github.vkostin.xlss.sheet;

import io.github.vkostin.xlss.styles.DefaultColors;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

import static io.github.vkostin.xlss.styles.StylesHelper.setBgColor;
import static io.github.vkostin.xlss.styles.StylesHelper.setNoBorder;

public final class SheetHelpers {
    public static void setDefaultSheetStyle(XSSFWorkbook workbook, Sheet sheet, Function<DefaultStyleRendererParams, XSSFCellStyle> createDefaultStyle) {
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

    public static Cell getOrCreateCell(XSSFRow destRow, int colIdx) {
        Cell cell = destRow.getCell(colIdx);
        if (cell == null) {
            cell = destRow.createCell(colIdx);
        }
        return cell;
    }

    public record DefaultStyleRendererParams(XSSFWorkbook workbook, Sheet sheet) {

    }

    public static Function<DefaultStyleRendererParams, XSSFCellStyle> getWhiteBgStyleGenerator() {
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

    public static List<Sheet> getNotHiddenSheets(XSSFWorkbook workbook) {
        List<Sheet> visibleSheets = new ArrayList<>();
        int numSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numSheets; i++) {
            if (!workbook.isSheetHidden(i) && !workbook.isSheetVeryHidden(i)) {
                visibleSheets.add(workbook.getSheetAt(i));
            }
        }
        return visibleSheets;
    }

    public static XSSFRow createRowIfNotExist(XSSFSheet sheet, int rowIdx) {
        if (sheet.getRow(rowIdx) == null) {
            sheet.createRow(rowIdx);
        }
        return sheet.getRow(rowIdx);
    }
}
