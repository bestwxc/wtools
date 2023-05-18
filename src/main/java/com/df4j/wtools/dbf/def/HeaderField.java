package com.df4j.wtools.dbf.def;

public enum HeaderField {

    FileVersion(0, 1, "Valid dBASE for DOS file; bits 0–2 indicate version number," +
            "bit 3 indicates the presence of a dBASE for DOS memo file, bits 4–6 indicate the presence of a SQL table," +
            "bit 7 indicates the presence of any memo file (either dBASE m PLUS or dBASE for DOS)"),
    LastUpdateDate(1, 3, "Date of last update; formatted as YYMMDD" +
            "(with YY being the number of years since 1900)"),
    RecordNumber(4, 4,"Number of records in the database file"),
    HeaderSize(8, 2,"Number of bytes in the header"),
    RecordSize(10, 2,"Number of bytes in the record"),
    Reserved12(12, 2,"Reserved; fill with 0"),
    IncompleteFlag(14, 1,"Flag indicating incomplete transaction"),
    EncryptionFlag(15, 1,"Encryption flag"),
    Reserved16(16, 12,"Reserved for dBASE for DOS in a multi-user environment"),
    MdxFlag(28, 1,"Production .mdx file flag; 1 if there is a production .mdx file, 0 if not"),
    LanguageDriverID(29, 1,"Language driver ID"),
    Reserved30(30, 2,"Reserved; fill with 0")
    ;

    private int offset = 0;

    private int size;

    private String description;

    private HeaderField(int offset, int size, String description) {
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
