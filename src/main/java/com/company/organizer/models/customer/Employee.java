package com.company.organizer.models.customer;

public class Employee extends DescriptionModel{

    private String name;
    private String title;
    private Integer employeeSeniority;
    private boolean finalRevenue;
    private String rm;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getRm() {
        return rm;
    }

    public void setRm(String rm) {
        this.rm = rm;
    }

    public boolean isFinalRevenue() {
        return finalRevenue;
    }

    public void setFinalRevenue(boolean finalRevenue) {
        this.finalRevenue = finalRevenue;
    }

    public Integer getEmployeeSeniority() {
        return employeeSeniority;
    }

    public void setEmployeeSeniority(Integer employeeSeniority) {
        this.employeeSeniority = employeeSeniority;
        }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

}

