package com.df4j.wtools.office.excel;

import com.df4j.wtools.base.utils.CloseUtils;
import com.df4j.wtools.base.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.extractor.XSSFEventBasedExcelExtractor;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class ExcelUtils {

    private static char START = 'A';
    private static char END = 'Z';
    private static int LEN = 'Z' - 'A' + 1;

    /**
     * 给定行号，得出从0开始的序号，比如A表示0，B表示1，AA表示26
     *
     * @param cellRef excel以字母表示的列序号
     * @return
     */
    public static int getCellIndex(String cellRef) {
        int index = 0;
        if (StringUtils.hasText(cellRef)) {
            int n = 0;
            for (int i = cellRef.length() - 1; i >= 0; i--) {
                char tmp = cellRef.charAt(i);
                if (tmp < START || tmp > END) {
                    continue;
                }
                int count = tmp - START + 1;
                for (int j = n; j > 0; j--) {
                    count = count * LEN;
                }
                index = index + count;
                n++;
            }
            index -= 1;
        }
        return index;
    }

    public static XSSFWorkbook createXSSFWorkbook() {
        return new XSSFWorkbook();
    }

    public static HSSFWorkbook createHSSFWorkbook() {
        return new HSSFWorkbook();
    }

    public static SXSSFWorkbook createSXSSFWorkbook(int rowAccessWindowSize) {
        // used to write, not read.
        return new SXSSFWorkbook(rowAccessWindowSize);
    }

    public static XSSFWorkbook readXSSFWorkbook(String path) {
        try {
            return new XSSFWorkbook(path);
        } catch (Exception e) {
            throw new RuntimeException("read xssf fail. path: " + path, e);
        }
    }

    public static HSSFWorkbook readHSSFWorkbook(String path) {
        try {
            InputStream inputStream = new FileInputStream(path);
            return new HSSFWorkbook(inputStream);
        } catch (Exception e) {
            throw new RuntimeException("read hssf fail. path: " + path, e);
        }
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

    public static WorkbookOperations hssf() {
        return workbook(createHSSFWorkbook());
    }

    public static WorkbookOperations xssf(String path) {
        return workbook(readHSSFWorkbook(path));
    }

    public static WorkbookOperations hssf(String path) {
        return workbook(readHSSFWorkbook(path));
    }

    public static WorkbookOperations sxssf() {
        return sxssf(100);
    }

    public static WorkbookOperations sxssf(int rowAccessWindowSize) {
        return workbook(createSXSSFWorkbook(rowAccessWindowSize));
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

    public static void handleBigExcel(String path, String sheetName, XSSFSheetXMLHandler.SheetContentsHandler sheetContentsHandler) {
        File file = new File(path);
        try {
            OPCPackage pkg = OPCPackage.open(file);
            XSSFEventBasedExcelExtractor excelExtractor = new XSSFEventBasedExcelExtractor(pkg);
            XSSFReader reader = new XSSFReader(pkg);
            StylesTable stylesTable = reader.getStylesTable();
            ReadOnlySharedStringsTable stringsTable = new ReadOnlySharedStringsTable(pkg);
            XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            InputStream inputStream = null;
            while (iterator.hasNext()) {
                try {
                    inputStream = iterator.next();
                    String tmpSheetName = iterator.getSheetName();
                    if (!sheetName.equalsIgnoreCase(tmpSheetName)) {
                        continue;
                    }
                    excelExtractor.processSheet(sheetContentsHandler, stylesTable, null, stringsTable, inputStream);
                } finally {
                    CloseUtils.close(inputStream);
                }
            }
        } catch (Exception e) {
            throw new RuntimeException("handle excel faile, file: " + file.getAbsolutePath(), e);
        }
    }
}
