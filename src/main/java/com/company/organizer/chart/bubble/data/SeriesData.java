package com.company.organizer.chart.bubble.data;

import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

/**
 * Chart series data
 */
public class SeriesData {

    private String name;
    private String xValue;
    private String yValue;
    private String size;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getXValue() {
        return xValue;
    }

    public void setXValue(String xValue) {
        this.xValue = xValue;
    }

    public String getYValue() {
        return yValue;
    }

    public void setYValue(String yValue) {
        this.yValue = yValue;
    }

    public String getSize() {
        return size;
    }

    public void setSize(String size) {
        this.size = size;
    }

    public List<String> getAll() {
        List<String> allData = new LinkedList<>();
        allData.add(name);
        allData.add(xValue);
        allData.add(yValue);
        allData.add(size);
        return Collections.unmodifiableList(allData);
    }
}
