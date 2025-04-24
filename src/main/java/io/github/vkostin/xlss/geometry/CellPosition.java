package io.github.vkostin.xlss.geometry;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * CellPosition coordinates are 0 based
 */
@Getter
@AllArgsConstructor
public class CellPosition {
    int x;
    int y;

    public static CellPosition fromString(String position) {
        return PositionHelper.pos2CellPos(position);
    }

    public String toExcelString() {
        return PositionHelper.int2pos(this.x) + (this.y + 1);
    }

    public CellPosition substract(CellPosition pointer) {
        return new CellPosition(this.x - pointer.x, this.y - pointer.y);
    }

    public CellPosition withX(int x) {
        return new CellPosition(x, this.y);
    }

    public CellPosition withY(int y) {
        return new CellPosition(this.x, y);
    }

    public CellPosition plusX(int deltaX) {
        return new CellPosition(this.x + deltaX, this.y);
    }

    public CellPosition plusY(int deltaY) {
        return new CellPosition(this.x, this.y + deltaY);
    }

    public CellPosition minusX(int deltaX) {
        return new CellPosition(this.x - deltaX, this.y);
    }

    public CellPosition minusY(int deltaY) {
        return new CellPosition(this.x, this.y - deltaY);
    }

    public CellPositionPointer withFixedXY() {
        return new CellPositionPointer(this.x, this.y, true, true);
    }

    public CellPositionPointer withFixedX() {
        return new CellPositionPointer(this.x, this.y, true, false);
    }

    public CellPositionPointer withFixedY() {
        return new CellPositionPointer(this.x, this.y, false, true);
    }

    public CellPosition with(CellPositionDelta cellPositionDelta) {
        return new CellPosition(x + cellPositionDelta.getDx(), y + cellPositionDelta.getDy());
    }

    @Override
    public String toString() {
        return this.toExcelString();
    }

}
