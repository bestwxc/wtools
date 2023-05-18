package com.df4j.wtools.dbf.def;

public enum FieldDef {

    Name(0, 11, "Field name in ASCII (zero-filled)"),
    Type(11,1,"Field type. Allowed values: C, D, F, L, M, or N (see next table for meanings)"),
    Reserved12(12,4,"Reserved"),
    FieldLength(16,1,"Field length in binary (maximum 254 (0xFE))."),
    FieldCount(17,1,"Field decimal count in binary"),
    WorkAreaID(18,2,"Work area ID"),
    Example(20,1,"Example"),
    Reserved21(21,10,"Reserved"),
    MdxFlag(31,1,"Production MDX field flag; 1 if field has an index tag in the production MDX file, 0 if not")

    ;
    private int offset = 0;

    private int size;

    private String description;

    private FieldDef(int offset, int size, String description) {
        this.offset = offset;
        this.size = size;
        this.description = description;
    }

    public int getOffset() {
        return offset;
    }

    public int getSize() {
        return size;
    }
}
