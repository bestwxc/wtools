package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookOperationsImp implements WorkbookOperations {

    private Workbook workbook;

    private WorkbookOperationsImp(Workbook workbook) {
        this.workbook = workbook;
    }

    public static WorkbookOperationsImp of(Workbook workbook) {
        return new WorkbookOperationsImp(workbook);
    }

    @Override
    public Workbook getOriginWorkbook() {
        return this.workbook;
    }

    @Override
    public WorkbookOperations write(String path) {
        ExcelUtils.write(this.workbook, path);
        return this;
    }

    @Override
    public SheetOperations sheet(String sheetName) {
        Sheet sheet = ExcelUtils.goc(this.workbook, sheetName);
        return SheetOperationsImpl.of(this, sheet);
    }
}
