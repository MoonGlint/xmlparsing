package rowstyles;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class StatementRow {
    private Workbook workbook;
    private CellStyle style;
    public StatementRow(Workbook workbook) {
    this.workbook=workbook;
    workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());

        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());

    }
    public CellStyle getStyle() {
        return style;
    }


}
