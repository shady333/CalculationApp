package com.company.organizer.chart.bubble.data;

public class ChartSize {

    private final int firstColumnIndex;
    private final int firstRowIndex;
    private final int lastColumnIndex;
    private final int lastRowIndex;

    public int getFirstColumnIndex() {
        return firstColumnIndex;
    }

    public int getFirstRowIndex() {
        return firstRowIndex;
    }

    public int getLastColumnIndex() {
        return lastColumnIndex;
    }

    public int getLastRowIndex() {
        return lastRowIndex;
    }

    public static class Builder {
        // Required parameters
        private final int lastColumnIndex;
        private final int lastRowIndex;
        // Optional parameters - initialized to default values
        private int firstColumnIndex = 0;
        private int firstRowIndex = 0;

        public Builder(int lastColumnIndex, int lastRowIndex) {
            this.lastColumnIndex = lastColumnIndex;
            this.lastRowIndex = lastRowIndex;
        }

        public Builder firstColumnIndex(int val) {
            firstColumnIndex = val;
            return this;
        }

        public Builder firstRowIndex(int val) {
            firstRowIndex = val;
            return this;
        }

        public ChartSize build() {
            return new ChartSize(this);
        }
    }

    private ChartSize(Builder builder) {
        firstColumnIndex = builder.firstColumnIndex;
        firstRowIndex = builder.firstRowIndex;
        lastColumnIndex = builder.lastColumnIndex;
        lastRowIndex = builder.lastRowIndex;
    }

}
