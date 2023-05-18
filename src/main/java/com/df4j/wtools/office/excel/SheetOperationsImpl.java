package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class SheetOperationsImpl implements WorkbookOperations, SheetOperations {

    private WorkbookOperations workbookOperations;

    private Sheet sheet;

    private SheetOperationsImpl(WorkbookOperations workbookOperations, Sheet sheet) {
        this.workbookOperations = workbookOperations;
        this.sheet = sheet;
    }

    public static SheetOperations of(WorkbookOperations workbookOperations, Sheet sheet) {
        return new SheetOperationsImpl(workbookOperations, sheet);
    }

    @Override
    public Sheet getOriginSheet() {
        return this.sheet;
    }

    @Override
    public WorkbookOperations workbook() {
        return this.workbookOperations;
    }

    @Override
    public RowOperations row(int rowIndex) {
        Row row = ExcelUtils.goc(this.sheet, rowIndex);
        return RowOperationsImpl.of(this, row);
    }

    @Override
    public Workbook getOriginWorkbook() {
        return this.workbook().getOriginWorkbook();
    }

    @Override
    public WorkbookOperations write(String path) {
        return this.workbook().write(path);
    }

    @Override
    public SheetOperations sheet(String sheetName) {
        return this.workbook().sheet(sheetName);
    }
}
