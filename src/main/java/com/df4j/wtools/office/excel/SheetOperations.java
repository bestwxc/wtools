package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Sheet;

public interface SheetOperations {

    Sheet getOriginSheet();

    WorkbookOperations workbook();
    RowOperations row(int rowIndex);
}
