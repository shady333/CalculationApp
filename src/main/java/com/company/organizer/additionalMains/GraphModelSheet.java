package com.company.organizer.additionalMains;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.emp.StatusEmp;
import com.company.organizer.core.graphDataSheet.GraphDataSheet;
import com.company.organizer.core.managersSheet.ProjectPMFunctionality;
import com.company.organizer.core.rev.RevenueParser;
import com.company.organizer.core.utils.Utils;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.graph.GraphEmp;

import java.util.List;

public class GraphModelSheet {
    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + Utils.getExcelPath();

        RevenueParser parser = new RevenueParser();
        GraphDataSheet graphEmpList = new GraphDataSheet();
        ProjectPMFunctionality functionality = new ProjectPMFunctionality();
        StatusEmp statusEmp = new StatusEmp();

        List<Customers> customers = parser.getCustomerModel();

        statusEmp.assignTitleToEmployeeAndFixedRev(customers);
        statusEmp.setFixedRev(customers);


        functionality.setPM(customers);
        List<GraphEmp> list = functionality.getGraphEmpList();
        graphEmpList.writeSheetGraphs(list);
    }
}
