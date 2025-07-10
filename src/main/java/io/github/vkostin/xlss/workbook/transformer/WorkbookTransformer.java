package io.github.vkostin.xlss.workbook.transformer;

import io.github.vkostin.xlss.geometry.CellPosition;
import io.github.vkostin.xlss.geometry.CellRange;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class WorkbookTransformer {
    /**
     * Copies a block of cells from one sheet to another sheet in the same workbook.
     * <p>
     * The source and target sheets are assumed to be in the same workbook.
     * <p>
     * The source block is determined by the bounds of the source range, and the target block is
     * a translation of the source block to the target position.
     * <p>
     * The copying is done in such a way that the formatting of the source block is preserved.
     * <p>
     * The rows and columns of the target block are created if they do not exist.
     *
     * @param source the source range
     * @param sourceSheet the source sheet
     * @param target the target position
     * @param targetSheet the target sheet
     */
    public static void copyBlockInSameWorkbook(
            CellRange source,
            XSSFSheet sourceSheet,
            CellPosition target,
            XSSFSheet targetSheet
    ) {
        new BlockBetweenDifferentSheetsCopier(source, sourceSheet, target, targetSheet).copyData();
    }
}
