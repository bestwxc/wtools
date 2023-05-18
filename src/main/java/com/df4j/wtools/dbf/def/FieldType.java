package com.df4j.wtools.dbf.def;

public enum FieldType {

    Character('C', "Any ASCII text (padded with spaces up to the field's length)"),
    Date('D', "\tNumbers and a character to separate month, day, and year (stored internally as 8 digits in YYYYMMDD format)"),
    FloatingPoint('F', "-, ., 0–9 (right justified, padded with whitespaces)"),
    Logical('L', "Y, y, N, n, T, t, F, f, or ? (when uninitialized)"),
    Memo('M', "Any ASCII text (stored internally as 10 digits representing a .dbt block number, right justified, padded with whitespaces)"),
    Numeric('N', "-, ., 0–9 (right justified, padded with whitespaces)");

    private char type;
    private String description;

    private FieldType(char type, String description) {
        this.type = type;
        this.description = description;
    }
}
