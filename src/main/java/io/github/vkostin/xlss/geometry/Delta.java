package io.github.vkostin.xlss.geometry;

import lombok.RequiredArgsConstructor;
import lombok.Value;

@Value
@RequiredArgsConstructor
public class Delta {
    int deltaX;
    int deltaY;
    CellPosition initialFrom;
    CellPosition initialTo;

    public static Delta from(CellPosition source, CellPosition target) {
        return new Delta(target.getX() - source.getX(), target.getY() - source.getY(), source, target);
    }

    public int applyToSourceY(int rowNum) {
        return rowNum + deltaY;
    }
    public int applyToSourceX(int colNum) {
        return colNum + deltaX;
    }
}
