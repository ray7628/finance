package com.inesat;


import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

import javax.annotation.Resource;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
@EnableConfigurationProperties
@Log4j2
public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

    @Resource
    private Mappings mappings;

    @Bean
    public CommandLineRunner commandLineRunner() {
        return args -> {
            try {
                String outputFileName = "主营收入成本明细账-分项目.xls";
                log.info("start transform from file {}.", mappings.getFilename());
                try (HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(mappings.getFilename()))) {
                    try (HSSFWorkbook wbOutput = new HSSFWorkbook()) {
                        for (String project : mappings.getProjects()) {
                            List<List<String>> incomeValues = readExcel(wb, project, mappings.getIncomeSheet(), mappings.getIncomeCols());
                            writeExcel(wbOutput,"主营收入-"+project, mappings.getIncomeCols(), incomeValues);

                            List<List<String>> costValues = readExcel(wb, project, mappings.getCostSheet(), mappings.getCostCols());
                            writeExcel(wbOutput,"主营成本-"+project, mappings.getCostCols(), costValues);
                        }

                        File file = new File(outputFileName);
                        try (OutputStream fileOut = new FileOutputStream(file)) {
                            wbOutput.write(fileOut);   //将workbook写入文件流
                        }
                        log.info("file write to : {}", file.getAbsolutePath());
                    }


                }


            } catch (Exception e) {
                log.error(e);
                int b = System.in.read();
                log.debug("input byte {} ", b);
            }
        };
    }

    public List<List<String>> readExcel(HSSFWorkbook wb, String project, String sheetName, List<ColMapping> cols) throws IOException {
        HSSFSheet sheet = wb.getSheet(sheetName);
        int projectColIndex = updateMappingIndexByTitleRow(cols, sheet);
        if (projectColIndex == -1) {
            throw new RuntimeException("项目编号 column not found!");
        }
        List<List<String>> values = new ArrayList<>();

        for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
            HSSFRow row = sheet.getRow(i);
            if (row == null || row.getCell(1) == null || StringUtils.isBlank(row.getCell(1).toString())) {
                break;
            }
            HSSFCell projectNumberCell = row.getCell(projectColIndex);
            if (projectNumberCell != null) {
                projectNumberCell.setCellType(CellType.STRING);
            }
            String projectNumber = projectNumberCell == null ? "" : projectNumberCell.toString();
            if (!projectNumber.equalsIgnoreCase(project)) continue;

            List<String> rowValue = new ArrayList<>();
            cols.forEach(e -> {
                HSSFCell cell = null;
                if (e.getColIndex() != null) {
                    cell = row.getCell(e.getColIndex());
                }
                rowValue.add(cell == null ? "" : cell.toString());
            });
            values.add(rowValue);
        }
        return values;
    }

    private int updateMappingIndexByTitleRow(List<ColMapping> cols, HSSFSheet sheet) {
        int result = -1;
        HSSFRow titleRow = sheet.getRow(0);
        if (titleRow != null) {
            List<String> titleRowValue = new ArrayList<>();
            for (int i = 0; i < titleRow.getLastCellNum(); i++) {
                HSSFCell cell = titleRow.getCell(i);
                String titleStr = (cell == null ? null : cell.toString());
                if ("项目编号".equals(titleStr)) result = i;
                titleRowValue.add(titleStr);
            }
            cols.forEach(e -> {
                if (StringUtils.isNotBlank(e.getFrom())) {
                    int index = titleRowValue.indexOf(e.getFrom());
                    if (index > -1) {
                        e.setColIndex(index);
                    }
                }
            });
        }

        return result;
    }

    public void writeExcel(Workbook wb, String sheetName,List<ColMapping> cols, List<List<String>> values) throws IOException {
        CreationHelper createHelper = wb.getCreationHelper();  //创建帮助工具

        /*创建表单*/
        Sheet sheet = wb.createSheet(sheetName);

        //设置字体
        Font headFont = wb.createFont();
        headFont.setFontHeightInPoints((short) 12);
        headFont.setFontName("宋体");

        /*设置数据单元格格式*/
        CellStyle dataStyle = wb.createCellStyle();
        dataStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setAlignment(HorizontalAlignment.CENTER);    //设置水平对齐方式
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        dataStyle.setFont(headFont);  //设置字体

        //设置头部单元格样式
        CellStyle headStyle = wb.createCellStyle();
        headStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        headStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderTop(BorderStyle.DOUBLE);
        headStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setAlignment(HorizontalAlignment.CENTER);    //设置水平对齐方式
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        headStyle.setFont(headFont);  //设置字体

        /*设置列宽度*/
        for (int i = 0; i <= cols.size(); i++) {
            sheet.setColumnWidth(i, 15 * 256);
        }

        int startRowIndex = 0;
        Row headRow = sheet.createRow(startRowIndex++);
        headRow.setHeight((short) 400);
        int col = 0;
         for (ColMapping colMapping : cols) {
            createTextCell(createHelper, headStyle, headRow, col++, colMapping.getTo());
        }

        for (List<String> rowValue : values) {
            Row valueRow = sheet.createRow(startRowIndex++);
            valueRow.setHeight((short) 300);
            col = 0;
            for (String value : rowValue) {
                createTextCell(createHelper, dataStyle, valueRow, col++, value);
            }
        }
    }

    private void createTextCell(CreationHelper createHelper, CellStyle cellStyle, Row row, int i, Object text) {
        Cell cell;
        cell = row.createCell(i);
        cell.setCellValue(createHelper.createRichTextString(text == null ? "" : text.toString()));
        cell.setCellStyle(cellStyle);
    }

}
