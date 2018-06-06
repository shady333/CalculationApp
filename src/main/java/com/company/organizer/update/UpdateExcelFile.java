package com.company.organizer.update;


import com.company.organizer.exception.InvalidColumnHeaderException;
import com.company.organizer.core.utils.ExcelFileUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class UpdateExcelFile {
    static Logger LOG = Logger.getLogger(UpdateExcelFile.class.getName());

    protected Workbook workbook;
    protected String fileName;

    public UpdateExcelFile(String fileName) {
        this.fileName = fileName;
    }

    public UpdateExcelFile(File file) {
        this.fileName = file.getPath();
    }

    private static int findColumn(Sheet sheet, String columnHeader) {
        int columnNumber = -1;
        Row row = sheet.getRow(0);
        for (Cell currentCell : row) {
            if (currentCell.getStringCellValue().equalsIgnoreCase(columnHeader)) {
                columnNumber = currentCell.getColumnIndex();
            }
        }
        return columnNumber;
    }

//    public UpdateExcelFile openFile() {
//        try {
//            InputStream inputStream;
//            inputStream = new FileInputStream(fileName);
//            workbook = WorkbookFactory.create(inputStream);
//            workbook.setMissingCellPolicy(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
//            inputStream.close();
//        } catch (InvalidFormatException e) {
//            LOG.error("Invalid file format");
//        } catch (FileNotFoundException e) {
//            LOG.error("File: " + fileName + " was not found");
//        } catch (IOException e) {
//            LOG.error("Error in IO stream");
//        }
//        if (workbook == null) {
//            throw new NullPointerException("Workbook wasn't instantiated properly");
//        }
//        return this;
//    }
//
    public UpdateExcelFile saveChangesToFile() {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(fileName);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            LOG.error("File: " + fileName + " was not found");
        } catch (IOException e) {
            LOG.error("Error in IO stream");
        }
        return this;
    }

    public UpdateExcelFile updateCellByIndex(int rowIndex, int cellIndex, String newCellValue) {
        Sheet sheet = workbook.getSheetAt(0);
        editCell(sheet, rowIndex, cellIndex, newCellValue);
        return this;
    }

    public UpdateExcelFile updateCellByColumnHeader(String columnHeader, int rowIndex, String newCellValue) {
        Sheet sheet = workbook.getSheetAt(0);
        int columnNumber = findColumn(sheet, columnHeader);
        if (columnNumber != -1) {
            editCell(sheet, rowIndex, columnNumber, newCellValue);
        } else throw new InvalidColumnHeaderException("No header '" + columnHeader + "' was found in file");

        return this;
    }

    public UpdateExcelFile updateColumnByHeader(String columnHeader, String newCellValue) {
        Sheet sheet = workbook.getSheetAt(0);
        int columnIndex = findColumn(sheet, columnHeader);
        if (columnIndex != -1) {
            for (Row currentRow : sheet) {
                if (currentRow.getRowNum() == 0)
                    continue;
                editCell(sheet, currentRow.getRowNum(), columnIndex, newCellValue);
            }
        } else throw new InvalidColumnHeaderException("No header '" + columnHeader + "' was found in file");

        return this;
    }

    public UpdateExcelFile updateColumnByIndex(int columnIndex, String newCellValue) {
        Sheet sheet = workbook.getSheetAt(0);
        int rowCount = ExcelFileUtils.getActualRowCount(sheet, 1);
        for (int currentRow = 1; currentRow <= rowCount; currentRow++) {
            editCell(sheet, currentRow, columnIndex, newCellValue);
        }
        return this;
    }

    public UpdateExcelFile updateRowByIndex(int rowIndex, String newCellValue) {
        Sheet sheet = workbook.getSheetAt(0);
        for (int cellIndex = 0; cellIndex < ExcelFileUtils.getActualCellCount(sheet); cellIndex++) {
            editCell(sheet, rowIndex, cellIndex, newCellValue);
        }
        return this;
    }

    private Cell editCell(Sheet sheet, int rowIndex, int cellIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        Cell cell = row.getCell(cellIndex);
        if (cell == null) {
            cell = row.createCell(cellIndex);
            cell.setCellValue(value);
        } else {
            cell.setCellValue(value);
        }
        return cell;
    }

    public Sheet getSheet(String sheetName) {
        return workbook.getSheet(sheetName);
    }


}
