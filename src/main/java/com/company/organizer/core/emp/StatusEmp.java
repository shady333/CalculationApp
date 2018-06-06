package com.company.organizer.core.emp;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.customer.Employee;
import com.company.organizer.models.rm.FullEmployee;
import org.apache.commons.collections4.CollectionUtils;

import java.util.*;

public class StatusEmp {
    private int number;
    private List<FullEmployee> empList = new EmployeeBhv().getEmployeeList();

    public String getSeniority(String value) {
        String normValue;
        if (value.equals("Junior Software Test Automation Engineer")) {
            normValue = CommonConst.JUNIOR;
            number = 1;
        } else if (value.equals("Software Test Automation Engineer")) {
            normValue = CommonConst.INTERMEDIATE;
            number = 2;
        } else if (value.equals("Senior Software Test Automation Engineer") || value.equals("Senior Software Testing Engineer")) {
            normValue = CommonConst.SENIOR;
            number = 3;
        } else if (value.equals("Lead Software Test Automation Engineer") || value.equals("Software Engineering Team Leader")) {
            normValue = CommonConst.LEAD;
            number = 4;
        } else if (value.equals("Software Engineering Manager")) {
            normValue = CommonConst.ABOVE_LEAD;
            number = 5;
        } else {
            number = 0;
            normValue = "doesn't found employee";
        }
        return normValue;
    }


    public List<FullEmployee> getBenchList(List<Customers> customers) {
        List<FullEmployee> benchList = new ArrayList<>();
        List<String> empNameProductionList = new ArrayList<>();
        List<String> empAllList = new ArrayList<>();

        customers.forEach(customer ->
                customer.getStreamsList().forEach(
                        streams -> streams.getEmployeesList().forEach(employee -> {
                            empNameProductionList.add(employee.getName());
                        })));

        empList.forEach(emp -> empAllList.add(emp.getName()));

        List<String> benchNameList = new ArrayList<>(CollectionUtils.subtract(empAllList, empNameProductionList));

        empList.forEach(emp -> benchNameList.forEach(s -> {
            if (s.equals(emp.getName()))
                benchList.add(emp);
        }));

        return benchList;
    }
//    public String findPersonTitle(String name, FullEmployee emp) {
//        for (FullEmployee employee : empList) {
//            if (name.equals(employee.getName())) {
//                return getSeniority(employee.getTitle());
//            }
//        }
//        if (name.equals(DOESN_T_FOUND_EMPLOYEE)) {
//            benchList.add(emp);
//        }
//        return DOESN_T_FOUND_EMPLOYEE;
//    }

    public String findPersonTitle(String name) {
        for (FullEmployee employee : empList) {
            if (name.equals(employee.getName())) {
                return getSeniority(employee.getTitle());
            }
        }
        return CommonConst.DOESN_T_FOUND_EMPLOYEE;
    }


    public String findPersonRM(String name) {
        for (FullEmployee employee : empList) {
            if (name.equals(employee.getName())) {
                return employee.getRm();
            }
        }
        return CommonConst.DOESN_T_FOUND_RM;
    }


    public void assignTitleToEmployeeAndFixedRev(List<Customers> customers) {
        for (Customers customer : customers) {
            for (int j = 0; j < customer.getStreamsList().size(); j++) {
                for (int k = 0; k < customer.getStreamsList().get(j).getEmployeesList().size(); k++) {
                    String person = customer.getStreamsList().get(j).getEmployeesList().get(k).getRowLabels();
                    if (!findPersonTitle(person).isEmpty()) {
                        String titleResult = findPersonTitle(person);
                        String RM = findPersonRM(person);
                        customer.getStreamsList().get(j).getEmployeesList().get(k).setName(person);
                        customer.getStreamsList().get(j).getEmployeesList().get(k).setTitle(titleResult);
                        customer.getStreamsList().get(j).getEmployeesList().get(k).setRm(RM);
                        customer.getStreamsList().get(j).getEmployeesList().get(k).setEmployeeSeniority(number);
                    }
                }
            }
        }
    }


    public void setFixedRev(List<Customers> customers) {
        for (Customers customer : customers) {
            for (int j = 0; j < customer.getStreamsList().size(); j++) {

                for (int k = 0; k < customer.getStreamsList().get(j).getEmployeesList().size(); k++) {

                    Double reporterHourse = Double.valueOf(customer.getStreamsList().get(j).getEmployeesList().get(k).getReportedHours());
                    String seniorityLevel = customer.getStreamsList().get(j).getEmployeesList().get(k).getTitle();
                    Double revenue = Double.valueOf(customer.getStreamsList().get(j).getEmployeesList().get(k).getRevenue());

                    try {

                        Double reporterHourse1 = Double.valueOf(customer.getStreamsList().get(j).getEmployeesList().get(k + 1).getReportedHours());
                        String seniorityLevel1 = customer.getStreamsList().get(j).getEmployeesList().get(k + 1).getTitle();
                        Double revenue1 = Double.valueOf(customer.getStreamsList().get(j).getEmployeesList().get(k + 1).getRevenue());


                        if (!reporterHourse.equals(reporterHourse1) && seniorityLevel.equals(seniorityLevel1) && revenue.equals(revenue1)) {
                            Employee employee = customer.getStreamsList().get(j).getEmployeesList().get(k);
                            employee.setFinalRevenue(true);

                            Employee employee1 = customer.getStreamsList().get(j).getEmployeesList().get(k + 1);
                            employee1.setFinalRevenue(true);
                        }
                    } catch (IndexOutOfBoundsException exc) {
                        exc.getStackTrace();
                    }
                }
            }
        }
    }

}
