package com.company.organizer.chart.bubble.data;

import com.company.organizer.chart.ChartData;

import java.util.Objects;

public class BubbleChartData implements ChartData {

    private final SourceData sourceData;
    private final ChartSize chartSize;

    public BubbleChartData(SourceData sourceData, ChartSize chartSize) {
        this.chartSize = Objects.requireNonNull(chartSize);
        this.sourceData = Objects.requireNonNull(sourceData);
    }

    @Override
    public ChartSize getChartSize() {
        return chartSize;
    }

    @Override
    public SourceData getSourceData() {
        return sourceData;
    }

}
