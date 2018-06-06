package com.company.organizer.additionalMains;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.rev.RevenueParser;
import com.company.organizer.core.utils.Utils;

public class ChartPoiRemoveSheetMain {
    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + Utils.getExcelPath();
        RevenueParser parser = new RevenueParser();
        parser.deleteEvaluationWarning();
    }
}
