package com.company.organizer.models.customer;

public class DescriptionModel {
    String rowLabels;
    String reportedHours;
    String revenue;
    String effectiveRate;



    public boolean isRowEmpty() {
        return getRowLabels().isEmpty() && getRevenue().isEmpty() && getReportedHours().isEmpty() &&
                getEffectiveRate().isEmpty();
    }

    public boolean isTestCaseCellEmpty() {
        return getRowLabels().isEmpty();
    }

    @Override
    public String toString() {
        return "DescriptionModel{" +
                "rowLabels='" + rowLabels + '\'' +
                ", reportedHours='" + reportedHours + '\'' +
                ", revenue='" + revenue + '\'' +
                ", effectiveRate='" + effectiveRate + '\'' +
                '}';
    }

    public String getRowLabels() {
        return rowLabels;
    }

    public void setRowLabels(String rowLabels) {
        this.rowLabels = rowLabels;
    }

    public String getReportedHours() {
        return reportedHours;
    }

    public void setReportedHours(String reportedHours) {
        this.reportedHours = reportedHours;
    }

    public String getRevenue() {
        return revenue;
    }

    public void setRevenue(String revenue) {
        this.revenue = revenue;
    }

    public String getEffectiveRate() {
        return effectiveRate;
    }

    public void setEffectiveRate(String effectiveRate) {
        this.effectiveRate = effectiveRate;
    }
}
