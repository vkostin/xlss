package io.github.vkostin.xlss.geometry;

import java.util.ArrayList;
import java.util.List;

public class CellRangeListHelper {

    public static List<CellRange> unite(List<CellRange> a, List<CellRange> b) {
        if (a == null && b == null) {
            return List.of();
        }
        if (a == null) {return b;}
        if (b == null) {return a;}
        List<CellRange> r = new ArrayList<>(a);
        r.addAll(b);
        return r.stream().distinct().toList();
    }

}
