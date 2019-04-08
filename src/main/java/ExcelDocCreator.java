import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.FileNotFoundException;

import java.io.FileOutputStream;
import java.io.IOException;

// TODO: 24.03.2019 start with testNg test 
// TODO: 24.03.2019 make table line visible
// TODO: 24.03.2019 plug in log4g 2 or latest 
// TODO: 24.03.2019 make some elements differ by color 
// TODO: 24.03.2019 use only write path relative not absolute(относительный не абсолютный) 
// TODO: 24.03.2019  ask user about name of excel file created 
// TODO: 24.03.2019 make this app instaleble 

public class ExcelDocCreator {
    private static String XLS_FILE_NAME = "src\\main\\resources\\sheet.xls";
    private HSSFWorkbook workbook = new HSSFWorkbook();
    private HSSFSheet sheet = workbook.createSheet("FirstSheet");

    private FileOutputStream fileOut;
    private HSSFRow rowNum;
    static int k = 0;

    public void createDefaultSheet() {


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
            sheet.autoSizeColumn(n);
        }

        System.out.println("Default table in excel file has been generated!");

    }

    // TODO: 24.03.2019 You should check the posibility of loosing some cells (rewriting it) 

    public void setCellData(String data, int rowIndex, int columnNum) {


        sheet.getRow(rowIndex).createCell(columnNum).setCellValue(data);
        sheet.getRow(rowIndex).setHeightInPoints(30);
        //sheet.autoSizeColumn(columnNum);

        //CellStyle cellStyle=workbook.createCellStyle();

        //System.out.println(" table in excel file has been generated!" + k++);
        System.out.println("rowIndex " + rowIndex + " , columnNum " + columnNum);


    }


    public void writeSheet() {
        try {
            fileOut = new FileOutputStream(XLS_FILE_NAME);
        } catch (FileNotFoundException e) {
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
