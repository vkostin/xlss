package io.github.vkostin.xlss.geometry;

import lombok.Getter;

@Getter
public class CellPositionPointer extends CellPosition {

    boolean xIsFixed;
    boolean yIsFixed;

    public CellPositionPointer(int x, int y, boolean xIsFixed, boolean yIsFixed) {
        super(x, y);
        this.xIsFixed = xIsFixed;
        this.yIsFixed = yIsFixed;
    }

    public CellPositionPointer(int x, int y) {
        this(x, y, false, false);
    }

    public CellPositionPointer withShiftFrom(CellPosition position) {
        int newX = xIsFixed ? x : (x + position.getX());
        int newY = yIsFixed ? y : (y + position.getY());
        return new CellPositionPointer(newX, newY, xIsFixed, yIsFixed);
    }

    @Override
    public String toExcelString() {
        return (xIsFixed ? "$" : "") + PositionHelper.int2pos(this.x) + (yIsFixed ? "$" : "") + (this.y + 1);
    }

}
