import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

// TODO: 24.03.2019 start with testNg test 
// TODO: 24.03.2019 make table line visible
// TODO: 24.03.2019 plug in log4g 
// TODO: 24.03.2019 make some elements differ by color 
// TODO: 24.03.2019 use only write path relative not absolute
// TODO: 24.03.2019  ask user about name of excel file created 
// TODO: 24.03.2019 make this app instalable

class ExcelDocCreator {
    private static String XLS_FILE_NAME = "src\\main\\resources\\sheet.xls";
    private HSSFWorkbook workbook = new HSSFWorkbook();
    private HSSFSheet sheet = workbook.createSheet("FirstSheet");
    private FileOutputStream fileOut;

    void createDefaultSheet() {


        CellStyle style = workbook.createCellStyle();

        style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());


        // foreground being the fill foreground not the font color
        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int i = 0; i < 800; i++) {
            sheet.createRow((short) i);

        }
        sheet.getRow(0).createCell(0).setCellValue("DATE");
        sheet.getRow(0).createCell(1).setCellValue("DEBET");
        sheet.getRow(0).createCell(2).setCellValue("CREDIT");
        sheet.getRow(0).createCell(3).setCellValue("BENEFICIAR");
        sheet.getRow(0).createCell(4).setCellValue("PAYER");
        sheet.getRow(0).createCell(5).setCellValue("GROUND");
        sheet.getRow(0).createCell(6).setCellValue("STATEMENT DATE");
        sheet.getRow(0).createCell(7).setCellValue("OPENING BALANCE");
        sheet.getRow(0).createCell(8).setCellValue("DEBET TURNOVER");
        sheet.getRow(0).createCell(9).setCellValue("CREDIT TURNOVER");
        sheet.getRow(0).createCell(10).setCellValue("CLOSING BALANCE");

        for (int n = 0; n <= 10; n++) {
            sheet.getRow(0).getCell(n).setCellStyle(style);
            //  sheet.autoSizeColumn(n);
        }

        System.out.println("Default table in excel file has been generated!");

    }

    // TODO: 24.03.2019 You should check the posebility of loosing some cells (rewriting it)

    void setCellData(String data, int rowIndex, int columnNum) {
        CellStyle style = workbook.createCellStyle();
        style.setWrapText(true);
       Row row = sheet.getRow(rowIndex);

//  the limit of cell length of 50 characters
        if (data.length() > 50) {
            for (int i = 1; i <= Math.abs(data.length() / 50); i++) {
                data = data.substring(0, i * 50) + "\n" + data.substring(i * 50);
            }

        }
        Cell cell = row.createCell(columnNum);
        row.setRowStyle(style);
        cell.setCellStyle(style);
        cell.setCellValue(data);
       // sheet.autoSizeColumn(0);





        System.out.println("rowIndex " + rowIndex + " , columnNum " + columnNum);


    }

    void setCellData(double data, int rowIndex, int columnNum) {

        Cell cell = sheet.getRow(rowIndex).createCell(columnNum);
        cell.setCellType(CellType.NUMERIC);
        cell.setCellValue(data);

        //System.out.println(" table in excel file has been generated!" + k++);
        System.out.println("rowIndex " + rowIndex + " , columnNum " + columnNum);


    }


    void writeSheet() {
        for (int n = 0; n <= 10; n++) {

            sheet.autoSizeColumn(n);
        }


        try {
            fileOut = new FileOutputStream(XLS_FILE_NAME);
        } catch (FileNotFoundException e) {
            System.err.println("!!!!!!!You should close your excel sheet and reboot class!!!!!!!!");
            e.printStackTrace();
        }
        try {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println(workbook);
        }
        try {
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}
