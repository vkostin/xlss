package io.github.vkostin.xlss.geometry;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public class Size {
    private final int width;
    private final int height;

    public Size(CellPosition tl, CellPosition br) {
        this.width = br.getX() - tl.getX();
        this.height = br.getY() - tl.getY();
    }

    @Override
    public String toString() {
        return width + "x" + height;
    }
}
