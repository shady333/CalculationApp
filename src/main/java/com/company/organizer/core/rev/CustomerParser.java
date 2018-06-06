package com.company.organizer.core.rev;

import com.company.organizer.commons.CommonConst;
import com.company.organizer.commons.NumberConstant;
import com.company.organizer.core.utils.Utils;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.core.utils.ExcelFileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ddf.EscherColorRef.SysIndexProcedure.THRESHOLD;

public class CustomerParser {
    private static int revenueRow = 10;
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.getSheet(CommonConst.REVENUE_SHEET);


    public List<String> getCustomerRowLabel() {
        List<String> customersRowLabel = new ArrayList<>();

        String streamName2 = "";
        do {
            if (isLastRevRow(revenueRow)) {
                break;
            }
            String rowLabel = getField(revenueRow, NumberConstant.ZERO.getNumber());
            String rowLabelPlusOne = getField(revenueRow + 1, NumberConstant.ZERO.getNumber());

            if (rowLabel.equals(THRESHOLD) || rowLabel.equals(CommonConst.THRESHOLD1) || rowLabel.equals(CommonConst.THRESHOLD2) || isRowRevEmpty(revenueRow)) {
                revenueRow++;
                continue;
            }
            if (Utils.isStream(rowLabelPlusOne)) {
                String streamName1 = Utils.getFirstPartOfStream(rowLabelPlusOne);
                if (!streamName1.equals(streamName2)) {
                    customersRowLabel.add(rowLabel);
                }
                streamName2 = Utils.getFirstPartOfStream(rowLabelPlusOne);
            }
            revenueRow++;
        } while (true);
        baseExcel.saveChangesToFile();
        return customersRowLabel;
    }

    public String getField(int rowNumber, Integer cellNumber) {
        Row row = sheet.getRow(rowNumber);
        try {
            return row.getCell(cellNumber).toString();
        } catch (NullPointerException ex) {
            return "";
        }
    }

    public boolean isRowRevEmpty(int rownum) {
        Row row = baseExcel.getSheet(CommonConst.REVENUE_SHEET).getRow(rownum);
        return ExcelFileUtils.isRowEmpty(row);
    }

    public boolean isLastRevRow(int rownum) {
        return isRowRevEmpty(rownum) & isRowRevEmpty(rownum + 1) & isRowRevEmpty(rownum + 2) & isRowRevEmpty(rownum + 3);
    }

}
