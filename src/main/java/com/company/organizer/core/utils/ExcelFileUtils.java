package com.company.organizer.core.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelFileUtils {

    public static boolean skip(final Row currentRow, int skip) {
        return currentRow.getRowNum() + 1 <= skip;
    }

    public static int getActualRowCount(Sheet sheet, int skip) {
        int rowsCount = 0;

        for (Row currentRow : sheet) {
            if (!isRowEmpty(currentRow)) {
                rowsCount++;
            }
        }
        return rowsCount - skip;
    }

    public static int getActualCellCount(Sheet sheet) {
        int maxCellCount = 0;

        for (Row currentRow : sheet) {
            if (!isRowEmpty(currentRow)) {
                int cellCount = 0;
                for (int cellIndex = 0; cellIndex < currentRow.getLastCellNum(); cellIndex++) {
                    cellCount++;
                    if (cellCount > maxCellCount) {
                        maxCellCount = cellCount;
                    }
                }
            }
        }
        return maxCellCount;
    }

    public static boolean isRowEmpty(Row row) {
        try {
            for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell != null && cell.getCellTypeEnum() != CellType.BLANK && StringUtils.isNotBlank(cell.toString()))
                    return false;
            }
        } catch (NullPointerException e) {
            return true;
        }
        return true;
    }
}
