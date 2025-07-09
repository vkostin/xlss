package io.github.vkostin.xlss.geometry;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CellRangeTest {

    @Test
    void normalized() {
        assertEquals("A1:B2", new CellRange(CellPosition.fromString("A1"), CellPosition.fromString("B2")).normalized().asExcelRangeAddress());
        assertThrows(IllegalArgumentException.class, (() -> assertEquals("A1:B2", new CellRange(CellPosition.fromString("B2"), CellPosition.fromString("A1")).normalized().asExcelRangeAddress())));
    }
}