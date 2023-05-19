package com.df4j.wtools.base.utils;

public class StringUtils {

    public static boolean hasLength(CharSequence s) {
        return s != null && s.length() > 0;
    }

    public static boolean hasText(CharSequence s) {
        return s != null && s.length() > 0 && containsText(s);
    }

    private static boolean containsText(CharSequence str) {
        int strLen = str.length();

        for (int i = 0; i < strLen; ++i) {
            if (!Character.isWhitespace(str.charAt(i))) {
                return true;
            }
        }
        return false;
    }
}
