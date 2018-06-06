package com.company.organizer.core.base;

import com.company.organizer.commons.CommonConst;
import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class BaseExcel {
    static Logger LOG = Logger.getLogger(BaseExcel.class.getName());

    protected Workbook workbook;
    protected String fileName;


    public BaseExcel(String fileName) {
        this.fileName = fileName;
    }

    public BaseExcel(File file) {
        this.fileName = file.getPath();
    }

    public BaseExcel openFile() {
        try {
            FileInputStream inputStream = new FileInputStream(fileName);
            workbook = WorkbookFactory.create(inputStream);
            inputStream.close();
        } catch (FileNotFoundException e) {
            LOG.error("File: " + CommonConst.REVENUE_PATH + " was not found");
        } catch (IOException e) {
            LOG.error("Error in IO stream");
        } catch (InvalidFormatException e) {
            LOG.error("Invalid file format");
        }
        if (workbook == null) {
            throw new NullPointerException("Workbook wasn't instantiated properly");
        }
        return this;
    }

    public BaseExcel saveChangesToFile() {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(CommonConst.REVENUE_PATH);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            LOG.error("File: " + CommonConst.REVENUE_PATH + " was not found");
        } catch (IOException e) {
            LOG.error("Error in IO stream");
        }
        return this;
    }

    public Sheet createSheet(String sheetName) {
        Sheet sheet;
        try {
            sheet = workbook.createSheet(sheetName);
        } catch (IllegalArgumentException exception) {
            workbook.removeSheetAt(workbook.getSheetIndex(sheetName));
            sheet = workbook.createSheet(sheetName);
        }

        return sheet;
    }

    public CreationHelper getCreationHelper() {
        return workbook.getCreationHelper();
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet(String sheetName) {
        return workbook.getSheet(sheetName);
    }

    public CellStyle createCellStyle() {
        return workbook.createCellStyle();
    }

    public FormulaEvaluator evaluateAllFormulaCells() {
        return workbook.getCreationHelper().createFormulaEvaluator();
    }

}
