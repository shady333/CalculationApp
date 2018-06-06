package com.company.organizer.commons;


import java.util.Arrays;
import java.util.List;

public class CommonConst {

    //    Headers
    public static final List<String> RESULT_COLUMN_HEADER_NAME = Arrays.asList("Project", "TA Name", "Reported Hours", "Base Hours", "Revenue", "Effective rate",
            "Revenue based on 152 h", "Lost revenue/0 rate", "Fixed Revenue", "Final Revenue", "Seniority Level",
            "Cost", "PM Real", "PM 152 HOURS without 0 rates", "Seniority per person", "Employee count",
            "Seniority per project", "Project PM", "Project PM 152");

    public static final List<String> RM_HEADER_NAME = Arrays.asList("RM","TA Name","Seniority Level","Emp Count","Average Seniority","Revenue","Revenue 152","Cost");
    public static final List<String> GRAPH_HEADER_NAME = Arrays.asList("Stream","Seniority per project","Project PM","Emp Count");
    //    Salary List
    public static final List<String> LEVELS_LIST = Arrays.asList("Level", "Salary", "Overhead", "Cost");

    public static final String CHART = "Chart";
    public static final String EVALUATION_WARNING = "Evaluation Warning";


    //  File & Path & Sheet constants
    public static final String FOLDER_PATH = ".";
    public static final String OUTPUT_DIRECTORY = "output";
    public static final String REVENUE_SHEET = "Revenue";
    public final static String EMPLOYEE_SHEET = "Sheet";
    public final static String CHART_PATH = "Chart/chart.xlsx";
    public static final String EMPLOYEE_PATH = "input/Employee.xlsx";
    public static final String MANAGERS_SHEET_NAME = "CreatedSheet";

    public static final String RM_SHEETS = "RMs";
    public static final String CHART_SHEET = "ChartSheet";

    public static final String JUNIOR = "Junior";
    public static final String INTERMEDIATE = "Intermediate";
    public static final String SENIOR = "Senior";
    public static final String LEAD = "Lead";
    public static final String ABOVE_LEAD = "AboveLead";
    // Excel Formulas
    public static final int BASE_HOURSE = 152;
    //    Kind of errors
    public static final String DOESN_T_FOUND_EMPLOYEE = "doesn't found employee";
        public static final String DOESN_T_FOUND_RM = "doesn't found RM";


    public static final String THRESHOLD = "Threshold is not exceeded, calculated based on adjusted rates";
    public static final String THRESHOLD1 = "Threshold is exceeded, Revenue per Employee model calculation";
    public static final String THRESHOLD2 = "Based on original rates";
    // The last word of list profiles
    public static final String GRAND_TOTAL = "Grand Total";
    public static String REVENUE_PATH = "PATH";

}
