package com.company.organizer.additionalMains;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.core.managersSheet.MargeFunctionality;
import com.company.organizer.core.utils.Utils;

public class MergeMain {
    public static void main(String[] args) {
        CommonConst.REVENUE_PATH = CommonConst.OUTPUT_DIRECTORY + "\\" + Utils.getExcelPath();
        new MargeFunctionality().mergeProjectsCells();
    }

}
