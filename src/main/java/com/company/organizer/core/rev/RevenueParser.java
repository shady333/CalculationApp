package com.company.organizer.core.rev;


import com.company.organizer.commons.CommonConst;
import com.company.organizer.commons.NumberConstant;
import com.company.organizer.core.utils.Utils;
import com.company.organizer.models.customer.Customers;
import com.company.organizer.models.customer.Employee;
import com.company.organizer.models.customer.Streams;
import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.core.utils.ExcelFileUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;


public class RevenueParser {
    public static int revenueRow = 10;
    private List<String> customerRowLabel = new CustomerParser().getCustomerRowLabel();
    private BaseExcel baseExcel = new BaseExcel(CommonConst.REVENUE_PATH).openFile();
    private Sheet sheet = baseExcel.getSheet(CommonConst.REVENUE_SHEET);


    public void createChartSheet() {
        baseExcel.createSheet(CommonConst.CHART);
        baseExcel.saveChangesToFile();
    }

    public void deleteEvaluationWarning() {

        int i = baseExcel.getWorkbook().getSheetIndex(CommonConst.EVALUATION_WARNING);
        if (i != -1) {
            baseExcel.getWorkbook().removeSheetAt(i);
        }
        baseExcel.saveChangesToFile();
    }

    public List<Customers> getCustomerModel() {
        List<Streams> streams = new ArrayList<>();
        List<Employee> employees = new ArrayList<>();
        List<Customers> customers = new ArrayList<>();
        Customers customerModel = null;
        Streams streamModel = null;
        Employee employeeModel;

        do {
            if (isLastRevRow(revenueRow)) {
                break;
            }

            String rowLabel = getField(revenueRow, NumberConstant.ZERO.getNumber());
            String rowLabelPlusOne = getField(revenueRow + 1, NumberConstant.ZERO.getNumber());

            if (rowLabel.equals(CommonConst.THRESHOLD) || rowLabel.equals(CommonConst.THRESHOLD1) || rowLabel.equals(CommonConst.THRESHOLD2) || isRowRevEmpty(revenueRow)) {
                revenueRow++;
                continue;
            }

            if (customerRowLabel.contains(rowLabel)) {
                customerModel = new Customers();
                customerModel.setRowLabels(getField(revenueRow, NumberConstant.ZERO.getNumber()));
                customerModel.setReportedHours(getField(revenueRow, NumberConstant.ONE.getNumber()));
                customerModel.setRevenue(getField(revenueRow, NumberConstant.TWO.getNumber()));
                customerModel.setEffectiveRate(getField(revenueRow, NumberConstant.THREE.getNumber()));

            } else if (Utils.isStream(rowLabel)) {
                streamModel = new Streams();

                streamModel.setRowLabels(getField(revenueRow, NumberConstant.ZERO.getNumber()));
                streamModel.setReportedHours(getField(revenueRow, NumberConstant.ONE.getNumber()));
                streamModel.setRevenue(getField(revenueRow, NumberConstant.TWO.getNumber()));
                streamModel.setEffectiveRate(getField(revenueRow, NumberConstant.THREE.getNumber()));

            } else {

                employeeModel = new Employee();
                employeeModel.setRowLabels(getField(revenueRow, NumberConstant.ZERO.getNumber()));
                employeeModel.setReportedHours(getField(revenueRow, NumberConstant.ONE.getNumber()));
                employeeModel.setRevenue(getField(revenueRow, NumberConstant.TWO.getNumber()));
                employeeModel.setEffectiveRate(getField(revenueRow, NumberConstant.THREE.getNumber()));
                employees.add(employeeModel);

                if (Utils.isStream(rowLabelPlusOne)) {
                    streamModel.setEmployeesList(new ArrayList<>(employees));
                    streams.add(streamModel);
                    employees.clear();

                } else if (customerRowLabel.contains(rowLabelPlusOne) || rowLabelPlusOne.equals(CommonConst.GRAND_TOTAL)) {
                    streamModel.setEmployeesList(new ArrayList<>(employees));
                    streams.add(streamModel);
                    employees.clear();

                    customerModel.setStreamsList(new ArrayList<>(streams));
                    customers.add(customerModel);
                    streams.clear();
                }
            }
            revenueRow++;
        } while (true);
        return customers;
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
        return isRowRevEmpty(rownum) && isRowRevEmpty(rownum + 1) && isRowRevEmpty(rownum + 2) &&
                isRowRevEmpty(rownum + 3) && isRowRevEmpty(rownum + 4) && isRowRevEmpty(rownum + 5);
    }

}
