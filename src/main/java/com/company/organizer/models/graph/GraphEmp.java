package com.company.organizer.models.graph;

public class GraphEmp {
    private String projectName;
    private Integer seniorityPerProject;
    private Double projectPM;
    private Integer empCount;

    public GraphEmp(String projectName, Integer seniorityPerProject, Double projectPM, Integer empCount) {
        this.projectName = projectName;
        this.seniorityPerProject = seniorityPerProject;
        this.projectPM = projectPM;
        this.empCount = empCount;
    }

    public String getProjectName() {
        return projectName;
    }

    public void setProjectName(String projectName) {
        this.projectName = projectName;
    }

    public Integer getSeniorityPerProject() {
        return seniorityPerProject;
    }

    public void setSeniorityPerProject(Integer seniorityPerProject) {
        this.seniorityPerProject = seniorityPerProject;
    }

    public Double getProjectPM() {
        return projectPM;
    }

    public void setProjectPM(Double projectPM) {
        this.projectPM = projectPM;
    }

    public Integer getEmpCount() {
        return empCount;
    }

    public void setEmpCount(Integer empCount) {
        this.empCount = empCount;
    }
}
