package com.company.organizer.core.managersSheet;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.commons.LevelConst;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.models.salaryTable.Position;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

public class BenchFunctionlity {
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.getSheet(CommonConst.MANAGERS_SHEET_NAME);

    public  List<Position> getLevelsRevenue() {
        List<Position> positionList = new ArrayList<>();

        Row benchRow = sheet.getRow(0);
        for (int i = 19; i < (CommonConst.LEVELS_LIST.size() + 19); i++) {
            Cell cell = benchRow.createCell(i);
            cell.setCellValue(CommonConst.LEVELS_LIST.get(i - 19));
        }

        LevelConst.DEFAULT_VALUE = Integer.parseInt(getField(1, 22));

        int LAST_ROW = 0;
        for (int i = 0, j = 2; i < 5; i++, j++) {

            Row benchRowForVal = sheet.getRow(j);

            Double salary  = Double.valueOf((getField(j, 20)));
            Double overhead  = Double.valueOf(getField(j, 21));

            positionList.add(new Position(getField(j, 19), salary.intValue(), overhead.intValue()));

            Cell cell22 = benchRowForVal.createCell(22);
            String formula = "U" + (j + 1) + "+V" + (j + 1);
            cell22.setCellFormula(formula);
            LAST_ROW = j;

        }
        LAST_ROW++;
//        Row averageRow = sheet.getRow(LAST_ROW);
        LevelConst.AVERAGE_RATE = Integer.parseInt(getField(LAST_ROW, 20));
        return positionList;
    }


    private String getField(int rowNumber, Integer cellNumber) {
        Row row = sheet.getRow(rowNumber);
        Cell cellValue = row.getCell(cellNumber);
        String string;
        try {
            switch (cellValue.getCellTypeEnum()) {
                case BOOLEAN:
                    string = String.valueOf(cellValue.getBooleanCellValue());
                    break;
                case NUMERIC:
                    string = String.valueOf(cellValue.getNumericCellValue());
                    break;
                case STRING:
                    string = cellValue.getStringCellValue();
                    break;
                case BLANK:
                    string = "Empty Cell";
                    break;
                case ERROR:
                    string = String.valueOf(cellValue.getErrorCellValue());
                    break;

                case FORMULA:
                    baseExcel.evaluateAllFormulaCells().evaluateInCell(cellValue);
                    string = String.valueOf(cellValue.getNumericCellValue());
                    break;
                default:
                    string = "Wow CELL have unrecognized type";
                    break;
            }
        } catch (NullPointerException e) {
            return String.valueOf(0);
        }

        return string;
    }
}
