package com.company.organizer.core.emp;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.commons.NumberConstant;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.core.utils.ExcelFileUtils;
import com.company.organizer.models.rm.FullEmployee;
import com.company.organizer.models.salaryTable.EmpTitle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;


public class EmployeeBhv {

    private BaseExcel baseExcel = new BaseExcel(CommonConst.EMPLOYEE_PATH).openFile();
    private Sheet sheet = baseExcel.getSheet(CommonConst.EMPLOYEE_SHEET);
    private List<FullEmployee> employeeList = new ArrayList<>();
    private HashMap<EmpTitle, String> npList = new HashMap<EmpTitle, String>();


    public List<FullEmployee> getEmployeeList() {
        int excelRow = 1;
        do {
            if (isLastEmpRow(excelRow)) {
                break;
            }
            String name = getField(excelRow, NumberConstant.ZERO.getNumber());
            String rm = getField(excelRow, NumberConstant.ONE.getNumber());
            String title = getField(excelRow, NumberConstant.TWO.getNumber());
            String primarySkill = getField(excelRow, NumberConstant.THREE.getNumber());

            FullEmployee employee = new FullEmployee(name, rm, title, primarySkill);
            employeeList.add(employee);

            excelRow++;
        }
        while (true);
        return employeeList;
    }

    public HashMap<EmpTitle, String> getBenchList() {
        int excelRow = 1;
        do {
            if (isLastEmpRow(excelRow)) {
                break;
            }
            String name = getField(excelRow, NumberConstant.ZERO.getNumber());
            String status = getField(excelRow, NumberConstant.TWO.getNumber());

            String np = getField(excelRow, 19);
            if (np.equals("1.0")) {
                npList.put(new EmpTitle(name, status), np);
            }
            excelRow++;
        }
        while (true);
        return npList;
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


    public boolean isEmptyEmpRow(int rownum) {
        Row row = sheet.getRow(rownum);
        return ExcelFileUtils.isRowEmpty(row);
    }

    public boolean isLastEmpRow(int rownum) {
        return isEmptyEmpRow(rownum) && isEmptyEmpRow(rownum + 1) && isEmptyEmpRow(rownum + 2);
    }
}
