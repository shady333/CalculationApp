package com.company.organizer;


import com.company.organizer.core.emp.StatusEmp;
import com.company.organizer.core.managersSheet.BenchFunctionlity;
import com.company.organizer.core.managersSheet.RevenueSheetWriter;
import com.company.organizer.core.rev.RevenueParser;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.salaryTable.Position;

import java.util.ArrayList;
import java.util.List;

import static com.company.organizer.commons.CommonConst.OUTPUT_DIRECTORY;
import static com.company.organizer.commons.CommonConst.REVENUE_PATH;
import static com.company.organizer.core.utils.Utils.copyFileUsingApacheCommonsIO;
import static com.company.organizer.core.utils.Utils.getExcelPath;

public class Main {
    public static void main(String[] args) {
        copyFileUsingApacheCommonsIO(getExcelPath(), OUTPUT_DIRECTORY);
        REVENUE_PATH = OUTPUT_DIRECTORY + "\\" + getExcelPath();

        List<Position> positionList = new ArrayList<>();

        try {
            BenchFunctionlity benchFunctionlity = new BenchFunctionlity();
            positionList = benchFunctionlity.getLevelsRevenue();
        } catch (Throwable ignored) {

        }

        RevenueParser parser = new RevenueParser();
        RevenueSheetWriter revenueSheetWriter = new RevenueSheetWriter();
        StatusEmp statusEmp = new StatusEmp();

        List<Customers> customers = parser.getCustomerModel();

        statusEmp.assignTitleToEmployeeAndFixedRev(customers);
        statusEmp.setFixedRev(customers);

        revenueSheetWriter.writeSheetEmp(customers);
        revenueSheetWriter.writeLevelsRevenue(positionList);
    }
}