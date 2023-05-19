package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Row;

public interface RowOperations extends SheetOperations {

    Row getOriginRow();

    SheetOperations sheet();

    CellOperations cell(int cellIndex);

    RowOperations fillRow(int startIndex, String[] values);
}
