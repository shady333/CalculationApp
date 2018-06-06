package com.company.organizer;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.emp.StatusEmp;
import com.company.organizer.core.managersSheet.ProjectPMFunctionality;
import com.company.organizer.core.rev.RevenueParser;
import com.company.organizer.models.customer.Customers;

import java.util.List;

import static com.company.organizer.core.utils.Utils.getExcelPath;

public class MainPM {

    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + getExcelPath();
        RevenueParser parser = new RevenueParser();
        ProjectPMFunctionality functionality = new ProjectPMFunctionality();
        StatusEmp statusEmp = new StatusEmp();


        List<Customers> customers = parser.getCustomerModel();

        statusEmp.assignTitleToEmployeeAndFixedRev(customers);
        statusEmp.setFixedRev(customers);


        functionality.setPM(customers);
    }
}
