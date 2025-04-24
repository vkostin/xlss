package io.github.vkostin.xlss.geometry;

import lombok.Getter;

import java.util.Objects;

@Getter
public class CellRange {
    final CellPosition topLeft;
    final CellPosition bottomRight;
    final int height;
    final int width;
    final String excelRangeAddress;

    public CellRange(CellPosition topLeft, int height, int width) {
        this.topLeft = topLeft;
        this.height = height;
        this.width = width;
        this.bottomRight = topLeft.plusY(height).plusX(width);
        this.excelRangeAddress = asExcelRangeAddress();
        if (this.height < 0 || this.width < 0) {
            throw new IllegalArgumentException("Invalid cell range : " + topLeft + "-" + bottomRight);
        }
    }

    public CellRange(CellPosition topLeft, CellPosition bottomRight) {
        this.topLeft = topLeft;
        this.bottomRight = bottomRight;
        this.height = bottomRight.getY() - topLeft.getY();
        this.width = bottomRight.getX() - topLeft.getX();
        this.excelRangeAddress = asExcelRangeAddress();
        if (this.height < 0 || this.width < 0) {
            throw new IllegalArgumentException("Invalid cell range : " + topLeft + "-" + bottomRight);
        }
    }

    public CellRange atTopLeft(CellPosition anchor) {
        return new CellRange(anchor, height, width);
    }

    public String asExcelRangeAddress() {
        return topLeft.toExcelString() + ":" + bottomRight.toExcelString();
    }

    public String asFixedExcelRangeAddress() {
        return new CellPositionPointer(topLeft.x, topLeft.y, true, true).toExcelString() + ":"
               + new CellPositionPointer(bottomRight.x, bottomRight.y, true, true).toExcelString();
    }

    @Override
    public boolean equals(final Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        return Objects.equals(excelRangeAddress, ((CellRange) obj).getExcelRangeAddress());
    }

    public CellRange normalized() {
        if (topLeft.getX() < bottomRight.getX() && topLeft.getY() > bottomRight.getY()) {
            return new CellRange(topLeft.withY(bottomRight.getY()), Math.abs(height), Math.abs(width));
        }
        if (topLeft.getX() > bottomRight.getX() && topLeft.getY() > bottomRight.getY()) {
            return new CellRange(bottomRight, Math.abs(height), Math.abs(width));
        }
        if (topLeft.getX() > bottomRight.getX() && topLeft.getY() < bottomRight.getY()) {
            return new CellRange(topLeft.withX(bottomRight.getX()),Math.abs(height), Math.abs(width));
        }
        return new CellRange(topLeft, Math.abs(height), Math.abs(width));
    }
}
