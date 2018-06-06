package com.company.organizer.core.managersSheet;

import com.company.organizer.core.base.BaseExcel;
import com.company.organizer.core.rev.RevenueParser;
import com.company.organizer.models.customer.Customers;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

import static com.company.organizer.commons.CommonConst.REVENUE_PATH;

public class MargeFunctionality {
    BaseExcel baseExcel = new BaseExcel(REVENUE_PATH).openFile();
    Sheet sheet = baseExcel.getSheet("CreatedSheet");
    private int BEGIN_ROW_CREATED_SHEET = 2;

    public void mergeProjectsCells() {
        RevenueParser parser = new RevenueParser();
        List<Customers> customers = parser.getCustomerModel();
        for (Customers customer : customers) {
            for (int j = 0; j < customer.getStreamsList().size(); j++) {
                int rowCount = BEGIN_ROW_CREATED_SHEET;
                for (int k = 0; k < customer.getStreamsList().get(j).getEmployeesList().size(); k++) {
                    sheet.getRow(1).getCell(0).setCellStyle(getProjectCellStyle());

                    Row row = getRow();
                    String string = row.getCell(0).getStringCellValue();
                    if (!string.equals("Bench")) {
                        row.getCell(0).setCellStyle(getProjectCellStyle());
                    }
                    BEGIN_ROW_CREATED_SHEET++;
                }
                int size = rowCount + (customer.getStreamsList().get(j).getEmployeesList().size() - 1);
                if (rowCount != size) {
                    sheet.addMergedRegion(CellRangeAddress.valueOf("A" + rowCount + ":A" + size));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("P" + rowCount + ":P" + size));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("Q" + rowCount + ":Q" + size));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("R" + rowCount + ":R" + size));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("S" + rowCount + ":S" + size));
                }

            }
        }

        baseExcel.saveChangesToFile();
    }

    public Row getRow() {
        return sheet.getRow(BEGIN_ROW_CREATED_SHEET);
    }

    public CellStyle getProjectCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        Font font = baseExcel.getWorkbook().createFont();
        font.setBold(true);

        style.setFont(font);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderTop(BorderStyle.THICK);
        return style;
    }

    public CellStyle getCountCellStyle() {
        CellStyle style = baseExcel.createCellStyle();

        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);

        Font font = baseExcel.getWorkbook().createFont();

        style.setFont(font);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderTop(BorderStyle.THICK);
        return style;
    }
}
