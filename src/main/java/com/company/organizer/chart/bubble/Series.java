package com.company.organizer.chart.bubble;

import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleSer;

public class Series {

    private CTBubbleSer series;

    public Series(CTBubbleSer bubbleChart) {
        this.series = bubbleChart;
    }

    public void setName(String name) {
        series.addNewTx().addNewStrRef().setF(name);
    }

    public void setColor(int rowNumber) {
        series.addNewIdx().setVal(rowNumber);
    }

    public void setXValue(String xValue) {
        series.addNewXVal().addNewStrRef().setF(xValue);
    }

    public void setYValue(String yValue) {
        series.addNewYVal().addNewNumRef().setF(yValue);
    }

    public void setSize(String reference) {
        series.addNewBubbleSize().addNewNumRef().setF(reference);
    }
}
