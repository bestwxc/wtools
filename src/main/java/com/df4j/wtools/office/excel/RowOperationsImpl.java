package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class RowOperationsImpl implements RowOperations, SheetOperations, WorkbookOperations {

    private SheetOperations sheetOperations;

    private Row row;

    private RowOperationsImpl(SheetOperations sheetOperations, Row row) {
        this.sheetOperations = sheetOperations;
        this.row = row;
    }

    public static RowOperationsImpl of(SheetOperations sheetOperations, Row row) {
        return new RowOperationsImpl(sheetOperations, row);
    }

    @Override
    public Row getOriginRow() {
        return this.row;
    }

    @Override
    public SheetOperations sheet() {
        return this.sheetOperations;
    }

    @Override
    public CellOperations cell(int cellIndex) {
        Cell cell = ExcelUtils.goc(this.row, cellIndex);
        return CellOperationsImpl.of(this, cell);
    }

    @Override
    public RowOperations fillRow(int startIndex, String[] values) {
        ExcelUtils.fillRow(this.row, startIndex, values);
        return this;
    }

    @Override
    public Sheet getOriginSheet() {
        return this.sheet().getOriginSheet();
    }

    @Override
    public WorkbookOperations workbook() {
        return this.sheetOperations.workbook();
    }

    @Override
    public RowOperations row(int rowIndex) {
        return this.sheetOperations.row(rowIndex);
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
