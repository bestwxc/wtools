package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Workbook;

public interface WorkbookOperations {

    Workbook getOriginWorkbook();

    WorkbookOperations write(String path);

    SheetOperations sheet(String sheetName);
}
