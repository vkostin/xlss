package io.github.vkostin.xlss.workbook.transformer;

import io.github.vkostin.xlss.geometry.CellPosition;
import io.github.vkostin.xlss.geometry.CellRange;
import io.github.vkostin.xlss.geometry.Delta;
import io.github.vkostin.xlss.sheet.SheetHelpers;
import lombok.AccessLevel;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Set;
import java.util.TreeSet;

import static io.github.vkostin.xlss.sheet.SheetHelpers.createRowIfNotExist;

@RequiredArgsConstructor(access = AccessLevel.PACKAGE)
class BlockBetweenDifferentSheetsCopier {
    private final CellRange sourceBlock;
    private final XSSFSheet sourceSheet;
    private final CellPosition target;
    private final XSSFSheet destSheet;

    private final Set<CellRange> usedMergedRegions = new TreeSet<>();

    public void copyData() {
        Delta delta = Delta.from(sourceBlock.getTopLeft(), target);
        int sourceWidth = sourceBlock.getWidth();

        sourceBlock.nonNullRowsStream(sourceSheet).forEach(row -> {
            int destY = delta.applyToSourceY(row.getRowNum());
            XSSFRow destRow = createRowIfNotExist(destSheet, destY);
            copyRow(row, destRow, delta, sourceWidth);
        });

        setColumnWidths(delta, sourceWidth);

    }

    private void setColumnWidths(Delta delta, int sourceWidth) {
        for (int i = sourceBlock.getTopLeft().getX(); i < sourceBlock.getBottomRight().getX() + 1; i++) {
            int destX = delta.applyToSourceX(i);
            destSheet.setColumnWidth(destX, sourceSheet.getColumnWidth(i));
            CellStyle srcColStyle = sourceSheet.getColumnStyle(i);
            if (srcColStyle != null)
            {
                destSheet.setDefaultColumnStyle(destX, srcColStyle);
            }
        }

    }

    private void copyRow(XSSFRow sourceRow, XSSFRow destRow, Delta delta, int sourceWidth) {
        if (sourceRow.isFormatted()) {
            CellStyle rowStyle = sourceRow.getRowStyle();
            destRow.setRowStyle(rowStyle);
        }

        short defaultHeight = sourceSheet.getDefaultRowHeight();
        if (sourceRow.getHeight() != defaultHeight) {
            destRow.setHeight(sourceRow.getHeight());
        }

        for (int col = sourceBlock.getTopLeft().getX(); col <= sourceBlock.getBottomRight().getX(); col++) {
            Cell sourceCell = sourceRow.getCell(col);
            Cell destCell = SheetHelpers.getOrCreateCell(destRow, delta.applyToSourceX(col));

            if (sourceCell != null) {
                copyCell(sourceCell, destCell);
                copyMergedRegion(sourceRow, sourceCell, delta);
            } else {
                destCell.setBlank();
            }
        }
    }

    private void copyMergedRegion(XSSFRow sourceRow, Cell sourceCell, Delta delta) {
        CellRangeAddress mergedRegion = getMergedRegion(sourceSheet, sourceRow.getRowNum(),
                (short) sourceCell.getColumnIndex());

        if (mergedRegion != null) {
            CellRangeAddress destMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
                    mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());

            CellRange range = CellRange.of(destMergedRegion);
            if (!usedMergedRegions.contains(range)) {
                usedMergedRegions.add(range);
                destSheet.addMergedRegion(range.asExcelRangeAddressObj());
            }
        }
    }

    private static CellRangeAddress getMergedRegion(XSSFSheet sheet, int rowNum, short cellNum) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowNum, cellNum)) {
                return merged;
            }
        }
        return null;
    }

    private void copyCell(Cell sourceCell, Cell destCell) {
        destCell.setCellStyle(sourceCell.getCellStyle());
        copyCellValue(sourceCell, destCell);
    }

    private static void copyCellValue(Cell sourceCell, Cell destCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                destCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(sourceCell.getNumericCellValue());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            case BOOLEAN:
                destCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case ERROR:
                destCell.setCellErrorValue(sourceCell.getErrorCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}
