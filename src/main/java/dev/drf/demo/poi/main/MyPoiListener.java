package dev.drf.demo.poi.main;

import dev.drf.demo.poi.core.PoiListener;

import java.util.HashMap;
import java.util.Map;

public class MyPoiListener implements PoiListener {
    private int sheetNum;
    private int currentRowNum;
    private Map<Integer, Map<Integer, String>> sheetValues;
    private Map<Integer, String> currentMap;

    public MyPoiListener() {
        this.sheetNum = -1;
        this.currentRowNum = -1;
        this.sheetValues = new HashMap<>();
        this.currentMap = new HashMap<>();
    }

    @Override
    public void nextSheet(int sheetNum) {
        this.sheetNum = sheetNum;
    }

    @Override
    public void nextRow(int rowNum) {
        this.currentRowNum = rowNum;
        this.sheetValues.putIfAbsent(rowNum, this.currentMap);
        this.currentMap = new HashMap<>();
    }

    @Override
    public void readCell(int cellNum, String value) {
        this.currentMap.putIfAbsent(cellNum, value);
    }

    public int getSheetNum() {
        return sheetNum;
    }

    public int getCurrentRowNum() {
        return currentRowNum;
    }

    public Map<Integer, Map<Integer, String>> getSheetValues() {
        return sheetValues;
    }
}
