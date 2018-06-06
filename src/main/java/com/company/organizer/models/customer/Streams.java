package com.company.organizer.models.customer;

import java.util.List;

public class Streams extends DescriptionModel {

    private List<Employee> employeesList;

    public List<Employee> getEmployeesList() {
        return employeesList;
    }

    public void setEmployeesList(List<Employee> employeesList) {
        this.employeesList = employeesList;
    }

    @Override
    public String toString() {
        return "Streams{" +
                "employeesList=" + employeesList +
                ", rowLabels='" + rowLabels + '\'' +
                ", reportedHours='" + reportedHours + '\'' +
                ", revenue='" + revenue + '\'' +
                ", effectiveRate='" + effectiveRate + '\'' +
                '}';
    }
}

