package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Sheet;

public interface SheetOperations extends WorkbookOperations{

    Sheet getOriginSheet();

    WorkbookOperations workbook();
    RowOperations row(int rowIndex);
}
