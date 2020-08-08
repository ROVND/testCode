package com.hongdy.code;

import cn.hutool.poi.excel.ExcelUtil;

import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        List<String> dayList = new ArrayList<>();

        for (int i = 0; i < 41; i++) {
            dayList.add(String.valueOf(i));
        }
        String value = getValue(dayList, 6, 7);
        System.out.println(value);
    }

    public static String getValue(List<String> dayList, Integer column, Integer row) {
        StringBuilder sb = new StringBuilder();
        sb.append("=MAX(");
        int size = dayList.size();
        int rowCount = dayList.size() / 13 + 1;
        List<String> startList = new ArrayList<>();
        List<String> endList = new ArrayList<>();
        for (int i = 0; i < rowCount * 2; i = i + 2) {
            String start = ExcelUtil.indexToColName(column) + (row + i);
            String end;
            int j = i / 2 + 1;
            if (j * 13 > size) {
                end = ExcelUtil.indexToColName(column + 13 - (j * 13 - size) - 1) + (row + i);
            } else {
                end = ExcelUtil.indexToColName(column + 13 - 1) + (row + i);
            }
            startList.add(start);
            endList.add(end);
        }

        for (int i = 0; i < startList.size(); i++) {
            sb.append(startList.get(i));
            sb.append(":");
            sb.append(endList.get(i));
            sb.append(",");
        }
        sb.deleteCharAt(sb.lastIndexOf(","));
        sb.append(")");

        sb.append("&\"shows\"&");
        sb.append("IF(");

        getP(sb, startList, endList);

        sb.append(">0,");

        sb.append("\"+\"&");

        getP(sb, startList, endList);

        sb.append("&\"shows\",\"\")&\"Test Show \"&TEXT(");
        String timeStart = ExcelUtil.indexToColName(column) + (row + 1);
        sb.append(timeStart);
        sb.append("-TIME(1,15,0),\"H:MM\")");

        return sb.toString();
    }

    private static void getP(StringBuilder sb, List<String> startList, List<String> endList) {
        for (int i = 0; i < startList.size(); i++) {
            sb.append("COUNTIF(");
            sb.append(startList.get(i));
            sb.append(":");
            sb.append(endList.get(i));
            sb.append(",");
            sb.append("\"p*\")");
            sb.append("+");
        }
        sb.deleteCharAt(sb.lastIndexOf("+"));
    }


}
