package io.github.vkostin.xlss.workbook.transformer;

import io.github.vkostin.xlss.geometry.CellPosition;
import io.github.vkostin.xlss.geometry.CellRange;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

public class WorkbookTransformerTest {
    @Test
    void testCopyBlockInSameWorkbook_copiesDataAndFormatting() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sourceSheet = workbook.createSheet("Source");
        XSSFSheet targetSheet = workbook.createSheet("Target");

        // Prepare source block: 2x2 block at (1,1)-(2,2)
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
        for (int r = 1; r <= 2; r++) {
            Row row = sourceSheet.createRow(r);
            for (int c = 1; c <= 2; c++) {
                Cell cell = row.createCell(c);
                cell.setCellValue("Val" + r + c);
                cell.setCellStyle(style);
            }
        }
        CellRange sourceRange = new CellRange(new CellPosition(1, 1), new CellPosition(2, 2));
        CellPosition targetPos = new CellPosition(5, 5);

        // Act
        WorkbookTransformer.copyBlockInSameWorkbook(sourceRange, sourceSheet, targetPos, targetSheet);

        // Assert: target block at (5,5)-(6,6) should match source block
        for (int dr = 0; dr < 2; dr++) {
            Row srcRow = sourceSheet.getRow(1 + dr);
            Row tgtRow = targetSheet.getRow(5 + dr);
            assertNotNull(tgtRow, "Target row should exist");
            for (int dc = 0; dc < 2; dc++) {
                Cell srcCell = srcRow.getCell(1 + dc);
                Cell tgtCell = tgtRow.getCell(5 + dc);
                assertNotNull(tgtCell, "Target cell should exist");
                assertEquals(srcCell.getStringCellValue(), tgtCell.getStringCellValue(), "Cell value");
                assertEquals(srcCell.getCellStyle().getWrapText(), tgtCell.getCellStyle().getWrapText(), "Cell style");
            }
        }
    }
}
