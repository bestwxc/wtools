package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.*;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public class CellOperationsImpl implements CellOperations, RowOperations, SheetOperations, WorkbookOperations {

    private RowOperations rowOperations;

    private Cell cell;

    private CellOperationsImpl(RowOperations rowOperations, Cell cell) {
        this.rowOperations = rowOperations;
        this.cell = cell;
    }


    public static CellOperationsImpl of(RowOperations rowOperations, Cell cell) {
        return new CellOperationsImpl(rowOperations, cell);
    }

    @Override
    public Row getOriginRow() {
        return null;
    }

    @Override
    public SheetOperations sheet() {
        return this.rowOperations.sheet();
    }

    @Override
    public CellOperations cell(int cellIndex) {
        return rowOperations.cell(cellIndex);
    }

    @Override
    public RowOperations fillRow(int startIndex, String[] values) {
        return this.rowOperations.fillRow(startIndex, values);
    }

    @Override
    public Sheet getOriginSheet() {
        return this.sheet().getOriginSheet();
    }

    @Override
    public WorkbookOperations workbook() {
        return this.sheet().workbook();
    }

    @Override
    public RowOperations row(int rowIndex) {
        return this.rowOperations;
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
        return this.rowOperations.sheet();
    }

    @Override
    public Cell getOriginCell() {
        return this.cell;
    }

    @Override
    public CellOperations setCellValue(String value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(double value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(Date value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(LocalDateTime value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(LocalDate value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(Calendar value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(RichTextString value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setCellValue(boolean value) {
        this.cell.setCellValue(value);
        return this;
    }

    @Override
    public CellOperations setAsActiveCell() {
        this.cell.setAsActiveCell();
        return this;
    }

    @Override
    public CellOperations setCellComment(Comment comment) {
        this.cell.setCellComment(comment);
        return this;
    }

    @Override
    public CellOperations setHyperlink(Hyperlink link) {
        this.cell.setHyperlink(link);
        return this;
    }

    @Override
    public CellOperations setBlank() {
        this.cell.setBlank();
        return this;
    }

    @Override
    public CellOperations setCellErrorValue(byte value) {
        this.cell.setCellErrorValue(value);
        return this;
    }

    @Override
    public CellOperations setCellFormula(String formula) {
        this.cell.setCellFormula(formula);
        return this;
    }

    @Override
    public CellOperations setCellStyle(CellStyle style) {
        this.cell.setCellStyle(style);
        return this;
    }
}
