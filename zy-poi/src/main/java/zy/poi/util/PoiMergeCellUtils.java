package zy.poi.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import zy.poi.config.MergeConfig;

import java.util.List;



/**
 * 纵向合并单元格工具类
 *
 * @author JueYue
 * @date 2015年6月21日 上午11:21:40
 */
public final class PoiMergeCellUtils {

    private PoiMergeCellUtils() {
    }

    /**
     * 纵向合并相同内容的单元格
     *
     * @param workbook
     * @param mergeConfig    合并所用参数对象
     */
    public static void mergeCells(Workbook workbook, MergeConfig mergeConfig) {
        boolean isMergeColumn=mergeConfig.isMergeColumn();
        List<List<int []>> merge = mergeConfig.getMergeRowParams();
        for (int z = 0, workbookLenght = workbook.getNumberOfSheets(); z < workbookLenght; z++) {
            List<int[]> params = merge.get(z);
            Sheet sheet = workbook.getSheetAt(z);
            for (int a = 0, paramsLenght = params.size(); a < paramsLenght; a++) {
                int[] startRow = params.get(a);
                /**
                 * [当前行,记录中断行,合并列]
                 */
                int[] rowCount = {startRow[0], startRow[0], startRow[1]};
                /**
                 * [当前所记录行的值,当前所记录列的值]
                 */
                String[] valueCount = {sheet.getRow(startRow[0]).getCell(0).getStringCellValue(), ""};
                for (int i = startRow[0], sheetLenght = sheet.getPhysicalNumberOfRows(); i < sheetLenght; i++) {
                    Row row = sheet.getRow(i);
                    Cell cell = row.getCell(0);
                    if (cell==null) {
                        mergerCell(sheet,rowCount,i);
                        continue;
                    }
                    String value = cell.getStringCellValue();
                    if (value != null && !value.equals("") && !valueCount[0].equals(value)) {
                        mergerCell(sheet,rowCount,i);
                        valueCount[0] = value;
                    }
                    if(isMergeColumn){
                        /**
                         * [当前列,记录列,合并行]
                         */
                        int [] columnCount = {1,0,i};
                        valueCount[1]=row.getCell(0).getStringCellValue();
                        for (int r = 1,rowLenght=row.getPhysicalNumberOfCells(); r <rowLenght ; r++) {
                            String columnValue=row.getCell(r).getStringCellValue();
                        }
                    }
                }
            }
        }
    }

    /**
     * 处理合并单元格的方法
     * @param sheet
     * @param rowCount
     * @param i
     */
    private static void mergerCell(Sheet sheet,int[] rowCount,int i){
        rowCount[1] = i - 1;
        if (i - rowCount[0] > 1) {
            try {
                sheet.addMergedRegion(new CellRangeAddress(rowCount[0], rowCount[1], rowCount[2], rowCount[2]));
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        rowCount[0] = i;
    }



}
