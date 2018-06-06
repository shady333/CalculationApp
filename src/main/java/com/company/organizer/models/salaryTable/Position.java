package com.company.organizer.models.salaryTable;

public class Position {
    private String level;
    private Integer salary;
    private Integer overhead;

    public Position(String level, Integer salary, Integer overhead) {
        this.level = level;
        this.salary = salary;
        this.overhead = overhead;
    }

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    public Integer getSalary() {
        return salary;
    }

    public void setSalary(Integer salary) {
        this.salary = salary;
    }

    public Integer getOverhead() {
        return overhead;
    }

    public void setOverhead(Integer overhead) {
        this.overhead = overhead;
    }
}

