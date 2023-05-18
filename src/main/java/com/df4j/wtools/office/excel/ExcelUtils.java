package com.df4j.wtools.office.excel;

import com.df4j.wtools.base.CloseUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import java.io.FileOutputStream;

public class ExcelUtils {

    public static XSSFWorkbook createXSSFWorkbook() {
        return new XSSFWorkbook();
    }


    /**
     * Get or create sheet by name.
     *
     * @param workbook
     * @param sheetName
     * @return
     */
    public static Sheet goc(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet != null ? sheet : workbook.createSheet(sheetName);
    }

    /**
     * Get or create row by index.
     *
     * @param sheet
     * @param rowIndex
     * @return
     */
    public static Row goc(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    /**
     * Get or create cell by index.
     *
     * @param row
     * @param cellIndex
     * @return
     */
    public static Cell goc(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex);
        return cell != null ? cell : row.createCell(cellIndex);
    }

    /**
     * from the startCellIndex cell, fill row with the columns
     *
     * @param row
     * @param startCellIndex
     * @param columns
     */
    public static void fillRow(Row row, int startCellIndex, String[] columns) {
        for (int i = 0; i < columns.length; i++) {
            Cell cell = goc(row, startCellIndex + i);
            if (StringUtils.hasText(columns[i])) {
                cell.setCellValue(columns[i]);
            } else {
                cell.setBlank();
            }
        }
    }

    public static WorkbookOperations workbook(Workbook workbook) {
        return WorkbookOperationsImp.of(workbook);
    }

    public static WorkbookOperations xssf() {
        return workbook(createXSSFWorkbook());
    }

    public static void write(Workbook workbook, String path) {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(path);
            workbook.write(out);
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            CloseUtils.close(out);
        }
    }
}
