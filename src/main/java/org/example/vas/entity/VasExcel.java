package org.example.vas.entity;

import lombok.Data;
import lombok.experimental.Accessors;
import org.example.vas.config.ExcelColumn;

/**
 * @author zhao
 * @time 2020/11/10 22:58
 */
@Data
@Accessors(chain = true)
public class VasExcel {
    @ExcelColumn(0)
    private String no;
    @ExcelColumn(1)
    private String number;
    @ExcelColumn(2)
    private String name;
    @ExcelColumn(3)
    private String date;
    @ExcelColumn(4)
    private String timeRange;
    @ExcelColumn(5)
    private String signInTime;
    @ExcelColumn(6)
    private String signOutTime;
}
