package com.company.organizer.models.rm;

public class FullEmployee {

    private String name;
    private String title;
    private Integer employeeSeniority;
    private String primarySkill;
    private String rm;

    public Integer getEmployeeSeniority() {
        return employeeSeniority;
    }

    public void setEmployeeSeniority(Integer employeeSeniority) {
        this.employeeSeniority = employeeSeniority;
    }

    public FullEmployee() {
    }

    public FullEmployee(String name, String rm, String title, String primarySkill) {
        this.name = name;
        this.rm = rm;
        this.title = title;
        this.primarySkill = primarySkill;
    }

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

    public String getTitle() {
            return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getPrimarySkill() {
        return primarySkill;
    }

    public void setPrimarySkill(String primarySkill) {
        this.primarySkill = primarySkill;
    }
}
