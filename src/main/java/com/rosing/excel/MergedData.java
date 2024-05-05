package com.rosing.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author luoxin
 */
@Data
@AllArgsConstructor
public class MergedData {

    private int firstRowIndex;

    private int firstColumnIndex;

    private int lastRowIndex;

    private int lastColumnIndex;

    private Cell realCell;



}
