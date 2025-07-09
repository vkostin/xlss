package io.github.vkostin.xlss.geometry;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class SizeTest {

    @Test
    void getDimensions() {
        assertEquals(2, new Size(1, 2).getHeight());

        assertEquals(3, new Size(new CellPosition(0, 0), new CellPosition(0, 3)).getHeight());
    }
}