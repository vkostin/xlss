package io.github.vkostin.xlss.geometry;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CellPositionTest {

    @Test
    void fromString() {
        assertEquals("A1", CellPosition.fromString("A1").toString());
        assertEquals(0, CellPosition.fromString("A1").getX());
        assertEquals(0, CellPosition.fromString("A1").getY());
        assertEquals("A2", CellPosition.fromString("A2").toString());
        assertEquals(0, CellPosition.fromString("A2").getX());
        assertEquals(1, CellPosition.fromString("A2").getY());
        assertEquals("B3", CellPosition.fromString("B3").toString());
        assertEquals(1, CellPosition.fromString("B2").getX());
        assertEquals(2, CellPosition.fromString("B3").getY());
    }
}