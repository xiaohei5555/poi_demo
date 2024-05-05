package com.rosing.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * @author luoxin
 */
public class MergedDataPool {

    private List<MergedData> list = new ArrayList<>();

    private Sheet sheet;

    public MergedDataPool(Sheet sheet){
        if(sheet != null){
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            mergedRegions.forEach(mergedRegion -> {
                int firstColumn = mergedRegion.getFirstColumn();
                int firstRow = mergedRegion.getFirstRow();
                int lastColumn = mergedRegion.getLastColumn();
                int lastRow = mergedRegion.getLastRow();
                Row row = sheet.getRow(firstRow);
                Cell cell = row.getCell(firstColumn);
                list.add(new MergedData(firstRow,firstColumn,lastRow,lastColumn,cell));
            });
        }
    }


    public boolean isMergedCell(Cell cell){
        if(cell == null){
            return false;
        }
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        for (MergedData mergedData:list) {
            int firstRowIndex = mergedData.getFirstRowIndex();
            int firstColumnIndex = mergedData.getFirstColumnIndex();
            int lastRowIndex = mergedData.getLastRowIndex();
            int lastColumnIndex = mergedData.getLastColumnIndex();
            if(rowIndex >= firstRowIndex && rowIndex <= lastRowIndex && columnIndex >= firstColumnIndex && columnIndex <= lastColumnIndex){
                return true;
            }
        }
        return false;
    }

    public Cell getRealCell(Cell cell){
        if(cell == null){
            return cell;
        }
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        for (MergedData mergedData:list) {
            int firstRowIndex = mergedData.getFirstRowIndex();
            int firstColumnIndex = mergedData.getFirstColumnIndex();
            int lastRowIndex = mergedData.getLastRowIndex();
            int lastColumnIndex = mergedData.getLastColumnIndex();
            if(rowIndex >= firstRowIndex && rowIndex <= lastRowIndex && columnIndex >= firstColumnIndex && columnIndex <= lastColumnIndex){
                return mergedData.getRealCell();
            }
        }
        return null;
    }

}
