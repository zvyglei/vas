package org.example.vas.config;

/**
 * @author zhao
 * @time 2020/12/6 19:22
 */
public enum ExcelParam {
    ROW_HEIGHT(35), COLUMN_WIDTH(12);

    private final int value;

    ExcelParam(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }
}