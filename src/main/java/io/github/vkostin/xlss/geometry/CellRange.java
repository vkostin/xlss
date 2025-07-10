package io.github.vkostin.xlss.geometry;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Objects;
import java.util.stream.IntStream;
import java.util.stream.Stream;

@Data
@Getter
public class CellRange implements Comparable<CellRange> {
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

    public CellRange(int firstColumn, int lastColumn, int firstRow, int lastRow) {
        this(new CellPosition(firstColumn, firstRow), lastColumn - firstColumn, lastRow - firstRow);
    }

    public static CellRange of(CellRangeAddress address) {
        return new CellRange(address.getFirstColumn(), address.getLastColumn(),
                             address.getFirstRow(), address.getLastRow());
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

    public String toString() {
        if (topLeft == null) {
            return "null";
        }
        return topLeft.toExcelString() + ":" + bottomRight.toExcelString();
    }

    public Stream<XSSFRow> nonNullRowsStream(XSSFSheet sourceSheet) {
        return IntStream.range(topLeft.getY(), bottomRight.getY() + 1)
                .mapToObj(sourceSheet::getRow)
                .filter(Objects::nonNull);
    }

    @Override
    public int compareTo(CellRange o) {
        CellPosition topLeft = this.getTopLeft();
        CellPosition oTopLeft = o.getTopLeft();
        if (topLeft.getY() < oTopLeft.getY()) {
            return -1;
        }
        if (topLeft.getY() > oTopLeft.getY()) {
            return 1;
        }
        if (topLeft.getX() < oTopLeft.getX()) {
            return -1;
        }
        if (topLeft.getX() > oTopLeft.getX()) {
            return 1;
        }
        // topLeft == oTopLeft
        if (this.getHeight() < o.getHeight()) {
            return -1;
        }
        if (this.getHeight() > o.getHeight()) {
            return 1;
        }
        return Integer.compare(this.getWidth(), o.getWidth());
    }

    public CellRangeAddress asExcelRangeAddressObj() {
        return new CellRangeAddress(topLeft.getY(), bottomRight.getY(), topLeft.getX(), bottomRight.getX());
    }
}
