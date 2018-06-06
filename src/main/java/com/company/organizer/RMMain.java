package com.company.organizer;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.rmSheet.RMSheet;
import com.company.organizer.core.utils.Utils;
import com.company.organizer.core.emp.EmployeeBhv;
import com.company.organizer.models.rm.FullEmployee;
import com.company.organizer.models.rm.RMPersonList;

import java.util.List;

public class RMMain {
    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + Utils.getExcelPath();

        RMSheet rmSheet = new RMSheet();
        List<FullEmployee> empList = new EmployeeBhv().getEmployeeList();

        List<RMPersonList> list = rmSheet.processAllEmployees(empList);
        rmSheet.writeSheetRMs(list);
    }
}
