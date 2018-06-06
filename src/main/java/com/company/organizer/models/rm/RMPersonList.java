package com.company.organizer.models.rm;

import java.util.List;

public class RMPersonList {
    private String rm;
    private List<FullEmployee> fullEmployees;


    public RMPersonList() {
    }

    public RMPersonList(String rm, List<FullEmployee> fullEmployees) {
        this.rm = rm;
        this.fullEmployees = fullEmployees;
    }

    public String getRm() {
        return rm;
    }

    public void setRm(String rm) {
        this.rm = rm;
    }

    public List<FullEmployee> getFullEmployees() {
        return fullEmployees;
    }

    public void setFullEmployees(List<FullEmployee> fullEmployees) {
        this.fullEmployees = fullEmployees;
    }
}
