package io.github.vkostin.xlss.geometry;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.lang.Integer.parseInt;

public final class PositionHelper {
    /**
     * Gets Excel string representation of the x anchorPosition in table. 0-based.
     * 0 gives A, 1 gives B, 25 gives Z. 26 gives AA. etc
     *
     * @param x
     * @return
     */
    public static String int2pos(int x) {
        if (x < 0) {
            throw new IllegalArgumentException("x must be >= 0");
        }
        int curValue = x;
        String acc = "";

        while (curValue >= 0) {
            int div = curValue / 26;
            int mod = curValue % 26;

            curValue = div - 1;
            acc = (char) ('A' + mod) + acc;
        }
        return acc;
    }

    final static String regex = "([A-Z]+)(\\d+)";

    public static String[] extract(String text) {
        final Pattern pattern = Pattern.compile(regex);
        final Matcher matcher = pattern.matcher(text);
        final String[] result = new String[2];
        if (matcher.find()) {
            result[0] = matcher.group(1);
            result[1] = matcher.group(2);
        }
        return result;
    }

    public static CellPosition pos2CellPos(String position) {
        String [] matches = extract(position);

        if (matches.length!= 2) {
            throw new IllegalArgumentException("Cannot parse " + position + " as Excel anchorPosition");
        }
        int x = str2int(matches[0]);
        int y = parseInt(matches[1]) -1;
        return new CellPosition(x, y);
    }

    public static int str2int(String str) {
        int val = 0;
        int power = 1;

        for (int i = str.length() - 1; i >= 0; i--) {
            int c = str.charAt(i) - 'A' + 1;
            val = val + (c * power);
            power = power * 26;
        }

        return val - 1;
    }
}
