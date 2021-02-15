package org.example.vas.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.example.vas.entity.VasExcel;
import org.example.vas.util.ExcelReadUtil;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author zhao
 * @time 2020/11/10 22:54
 */
public class GenerateExcel {
    public void generate(String path) {
        File folder = new File(path);
        if (!folder.exists()) {
            System.out.println("文件夹不存在！");
            return;
        }
        String month = "";
        List<File> files = Arrays.asList(folder.listFiles());
        if (files.size() > 0) {
            List<File> collect = files.stream().filter(x -> x.getName().endsWith(".xls") || x.getName().endsWith(".xlsx"))
                    .sorted(Comparator.comparing(x -> Integer.parseInt(x.getName().split("\\.")[1])))
                    .collect(Collectors.toList());

            month = collect.get(0).getName().split("\\.")[0];
            List<List<VasExcel>> dataList = new ArrayList<>();
            for (File file : collect) {
                Workbook workBook = ExcelReadUtil.getWorkBook(file);
                if(workBook != null){
                    Sheet sheet = workBook.getSheetAt(0);
                    int firstRowNum  = sheet.getFirstRowNum();
                    //获得当前sheet的结束行
                    int lastRowNum = sheet.getLastRowNum();

                    Row row0 = sheet.getRow(0);
                    Map headerMap = new HashMap();
                    for (Cell cell : row0) {
                        headerMap.put(cell.getColumnIndex(), cell.getRichStringCellValue());
                    }
                    //循环除了第一行的所有行
                    List<VasExcel> vasExcelList = new ArrayList<>();
                    for(int rowNum = firstRowNum+1;rowNum <= lastRowNum;rowNum++){
                        //获得当前行
                        Row row = sheet.getRow(rowNum);
                        if(row == null){
                            continue;
                        }
                        //获得当前行的开始列
                        int firstCellNum = row.getFirstCellNum();
                        //获得当前行的列数
                        int lastCellNum = row.getPhysicalNumberOfCells();

                        //循环当前行
                        VasExcel vasExcel = new VasExcel();
                        for (Cell cell : row) {
                            String key = headerMap.get(cell.getColumnIndex()).toString();
                            String val = ExcelReadUtil.getCellValue(cell);
                            switch (key) {
                                case "序号":
                                    vasExcel.setNo(val);
                                    break;
                                case "自定义编号":
                                    vasExcel.setNumber(val);
                                    break;
                                case "姓名":
                                    vasExcel.setName(val);
                                    break;
                                case "日期":
                                    vasExcel.setDate(val);
                                    break;
                                case "对应时段":
                                    vasExcel.setTimeRange(val);
                                    break;
                                case "签到时间":
                                    vasExcel.setSignInTime(val);
                                    break;
                                case "签退时间":
                                    vasExcel.setSignOutTime(val);
                                    break;
                            }
                        }
                        vasExcelList.add(vasExcel);
                    }
                    dataList.add(vasExcelList);
                }
                try {
                    workBook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            writeExcel(dataList, path, month);
        }
    }

    public void writeExcel(List<List<VasExcel>> dataList, String path, String month) {
        // 写法1
        String fileName = path + "/" + month + "月汇总-" + System.currentTimeMillis() + ".xlsx";

        String title = "朱家角规划资源所日考勤报表";
        String[] subTitles = {"序号","自定义编号","姓名","日期","对应时段","签到时间","签退时间"};
        try {
            HSSFWorkbook workBook = new HSSFWorkbook();
            HSSFSheet sheet = workBook.createSheet();
            HSSFCellStyle cellStyle = workBook.createCellStyle();
            HSSFCellStyle titleCellStyle = workBook.createCellStyle();

            HSSFFont font = workBook.createFont();
            font.setFontName("宋体");
            font.setFontHeightInPoints((short) 26);//设置字体大小

            sheet.setDefaultColumnWidth((short) 10);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
            cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
            cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
            cellStyle.setBorderTop(BorderStyle.THIN);//上边框
            cellStyle.setBorderRight(BorderStyle.THIN);//右边框


            titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
            titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
            titleCellStyle.setBorderBottom(BorderStyle.THIN); //下边框
            titleCellStyle.setBorderLeft(BorderStyle.THIN);//左边框
            titleCellStyle.setBorderTop(BorderStyle.THIN);//上边框
            titleCellStyle.setBorderRight(BorderStyle.THIN);//右边框
            titleCellStyle.setFont(font);

            int rowIndex = 0;
            for (List<VasExcel> ele : dataList) {
                Row titleRow = sheet.createRow(rowIndex);
                titleRow.setHeight((short) (35 * 20));
                Cell titleCell = titleRow.createCell(0);
                titleCell.setCellStyle(titleCellStyle);
                titleCell.setCellType(CellType.STRING);
                titleCell.setCellValue(title);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 6);
                sheet.addMergedRegion(cellRangeAddress);
                rowIndex += 1;

                Row subTitleRow = sheet.createRow(rowIndex);
                subTitleRow.setHeight((short) (35 * 20));
                for (int i = 0; i < subTitles.length; i++) {
                    Cell subTitleCell = subTitleRow.createCell(i);
                    subTitleCell.setCellStyle(cellStyle);
                    subTitleCell.setCellType(CellType.STRING);
                    subTitleCell.setCellValue(subTitles[i]);
                }
                rowIndex += 1;

                for (VasExcel vasExcel : ele) {
                    Row row = sheet.createRow(rowIndex);
                    row.setHeight((short) (35 * 20));
                    for (int i = 0; i < subTitles.length; i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellStyle(cellStyle);
                        cell.setCellType(CellType.STRING);
                        switch (i) {
                            case 0:
                                cell.setCellValue(vasExcel.getNo());
                                break;
                            case 1:
                                cell.setCellValue(vasExcel.getNumber());
                                break;
                            case 2:
                                cell.setCellValue(vasExcel.getName());
                                break;
                            case 3:
                                cell.setCellValue(vasExcel.getDate());
                                break;
                            case 4:
                                cell.setCellValue(vasExcel.getTimeRange());
                                break;
                            case 5:
                                cell.setCellValue(vasExcel.getSignInTime());
                                break;
                            case 6:
                                cell.setCellValue(vasExcel.getSignOutTime());
                                break;
                        }
                    }
                    rowIndex += 1;
                }

                // for (VasExcel vasExcel : ele) {
                //     Class<? extends VasExcel> vasExcelClass = vasExcel.getClass();
                //     Field[] fields = vasExcelClass.getDeclaredFields();
                //     for (Field field : fields) {
                //         boolean fieldHasAnno = field.isAnnotationPresent(ExcelColumn.class);
                //         if(fieldHasAnno){
                //             ExcelColumn fieldAnno = field.getAnnotation(ExcelColumn.class);
                //             int column = fieldAnno.value();
                //             if(!field.isAccessible()){
                //                 field.setAccessible(true);
                //             }
                //
                //             Row row = sheet.createRow(rowIndex);
                //             if (field.get(vasExcel) != null) {
                //                 String fieldValue = field.get(vasExcel).toString();
                //
                //                 Cell cell = row.createCell(column);
                //                 cell.setCellType(CellType.STRING);
                //                 cell.setCellValue(fieldValue);
                //             }
                //         }
                //     }
                //     rowIndex += 1;
                // }
            }

            FileOutputStream out = null;
            try {
                //创建文件
                File file = new File(fileName);
                file.createNewFile();
                out = new FileOutputStream(file);

                workBook.write(out);
                out.flush();
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    if (null != out) {
                        out.close();
                    }
                    if (null != workBook) {
                        workBook.close();
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
