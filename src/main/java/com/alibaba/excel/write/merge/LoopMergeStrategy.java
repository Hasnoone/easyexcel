package com.alibaba.excel.write.merge;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import com.alibaba.excel.metadata.property.LoopMergeProperty;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;

/**
 * The regions of the loop merge
 *
 * @author Jiaju Zhuang
 */
public class LoopMergeStrategy extends AbstractRowWriteHandler {
    /**
     * Each row
     */
    private int eachRow;
    /**
     * Extend column
     */
    private int columnExtend;
    /**
     * The number of the current column
     */
    private int columnIndex;


    private Row oldRow;

    private int dateRowCount;
    private int storeRowCount;
    private int brandRowCount;


    public LoopMergeStrategy(int eachRow, int columnIndex) {
        this(eachRow, 1, columnIndex);
    }

    public LoopMergeStrategy(int eachRow, int columnExtend, int columnIndex) {
        if (eachRow < 1) {
            throw new IllegalArgumentException("EachRows must be greater than 1");
        }
        if (columnExtend < 1) {
            throw new IllegalArgumentException("ColumnExtend must be greater than 1");
        }
        if (columnExtend == 1 && eachRow == 1) {
            throw new IllegalArgumentException("ColumnExtend or eachRows must be greater than 1");
        }
        if (columnIndex < 0) {
            throw new IllegalArgumentException("ColumnIndex must be greater than 0");
        }
        this.eachRow = eachRow;
        this.columnExtend = columnExtend;
        this.columnIndex = columnIndex;
    }

    public LoopMergeStrategy(LoopMergeProperty loopMergeProperty, Integer columnIndex) {
        this(loopMergeProperty.getEachRow(), loopMergeProperty.getColumnExtend(), columnIndex);
    }

    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row,
                                Integer relativeRowIndex, Boolean isHead) {
        if (isHead) {
            return;
        }

        if (null != oldRow) {
            Cell oldBrandCell = oldRow.getCell(0);
            String oldBrandVal = oldBrandCell.getStringCellValue();
            Cell oldStoreCell = oldRow.getCell(1);
            String oldStoreVal = oldStoreCell.getStringCellValue();
            Cell oldDateCell = oldRow.getCell(2);
            String oldDateVal = oldDateCell.getStringCellValue();

            Cell newBrandCell = row.getCell(0);
            String newBrandVal = newBrandCell.getStringCellValue();
            Cell newStoreCell = row.getCell(1);
            String newStoreVal = newStoreCell.getStringCellValue();
            Cell newDateCell = row.getCell(2);
            String newDateVal = newDateCell.getStringCellValue();

            int firstRow = row.getRowNum() -1;
            int lastRow = row.getRowNum()-1;

            if (oldBrandVal.equals(newBrandVal)) {
                brandRowCount++;
                if (oldStoreVal.equals(newStoreVal)) {
                    storeRowCount++;
                    if (oldDateVal.equals(newDateVal)) {
                        dateRowCount++;
                    }
                    else {
                        CellRangeAddress cellRangeAddress3 = new CellRangeAddress(firstRow - dateRowCount, lastRow,
                            2, 2);

                        if (firstRow - dateRowCount!=lastRow) {
                            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                        }
                        dateRowCount = 0;
                    }
                } else {
                    CellRangeAddress cellRangeAddress2 = new CellRangeAddress(firstRow - storeRowCount, lastRow,
                        1, 1);
                    CellRangeAddress cellRangeAddress3 = new CellRangeAddress(firstRow - dateRowCount, lastRow,
                        2, 2);

                    if (firstRow - dateRowCount!=lastRow) {
                        writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                    }
                    if (firstRow - storeRowCount!=lastRow) {
                        writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
                    }
                    dateRowCount = 0;
                    storeRowCount = 0;

                }
            } else {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow - brandRowCount, lastRow,
                    0, 0);
                CellRangeAddress cellRangeAddress3 = new CellRangeAddress(firstRow - dateRowCount, lastRow,
                    2, 2);
                CellRangeAddress cellRangeAddress2 = new CellRangeAddress(firstRow - storeRowCount, lastRow,
                    1, 1);
                if (firstRow - brandRowCount!=lastRow) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
                }
                writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
                if (firstRow - dateRowCount!=lastRow) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                }
                if (firstRow - storeRowCount!=lastRow) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
                }
                dateRowCount = 0;
                storeRowCount = 0;
                brandRowCount = 0;
            }
        }

        oldRow = row;



        //原方法
//        if (relativeRowIndex % eachRow == 0) {
//            CellRangeAddress cellRangeAddress = new CellRangeAddress(row.getRowNum(), row.getRowNum() + eachRow - 1,
//                columnIndex, columnIndex + columnExtend - 1);
//            CellRangeAddress cellRangeAddress2 = new CellRangeAddress(row.getRowNum(), row.getRowNum() + eachRow - 1,
//                1, columnIndex + 2 - 1);
//            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
//            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
//        }
    }

}
