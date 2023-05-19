package com.df4j.wtools.office.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

public interface CellOperations extends RowOperations{

    Cell getOriginCell();

    CellOperations setCellValue(String value);
    CellOperations setCellValue(double value);
    CellOperations setCellValue(Date value);
    CellOperations setCellValue(LocalDateTime value);
    CellOperations setCellValue(LocalDate value);
    CellOperations setCellValue(Calendar value);
    CellOperations setCellValue(RichTextString value);
    CellOperations setCellValue(boolean value);


}
