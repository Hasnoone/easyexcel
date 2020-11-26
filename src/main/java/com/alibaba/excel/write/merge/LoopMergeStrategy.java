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
    private int rowCount;

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

    public LoopMergeStrategy(int rowCount) {
        this.rowCount = rowCount;
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

        int rowNum = row.getRowNum();

        int rowNumber = rowNum - 1;


        if (1 != rowCount) {
            if (rowNum == rowCount) {
                dateRowCount++;
                storeRowCount++;
                brandRowCount++;
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowNum - brandRowCount, rowNum,
                    0, 0);
                CellRangeAddress cellRangeAddress3 = new CellRangeAddress(rowNum - dateRowCount, rowNum,
                    2, 2);
                CellRangeAddress cellRangeAddress2 = new CellRangeAddress(rowNum - storeRowCount, rowNum,
                    1, 1);
                writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
                writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
                return;
            }
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

            if (oldBrandVal.equals(newBrandVal)) {
                brandRowCount++;
                if (oldStoreVal.equals(newStoreVal)) {
                    storeRowCount++;
                    if (oldDateVal.equals(newDateVal)) {
                        dateRowCount++;
                    } else {
                        CellRangeAddress cellRangeAddress3 = new CellRangeAddress(rowNumber - dateRowCount, rowNumber,
                            2, 2);

                        if (rowNumber - dateRowCount != rowNumber) {
                            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                        }
                        dateRowCount = 0;
                    }
                } else {
                    CellRangeAddress cellRangeAddress2 = new CellRangeAddress(rowNumber - storeRowCount, rowNumber,
                        1, 1);
                    CellRangeAddress cellRangeAddress3 = new CellRangeAddress(rowNumber - dateRowCount, rowNumber,
                        2, 2);

                    if (rowNumber - dateRowCount != rowNumber) {
                        writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                    }
                    if (rowNumber - storeRowCount != rowNumber) {
                        writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
                    }
                    dateRowCount = 0;
                    storeRowCount = 0;

                }
            } else {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowNumber - brandRowCount, rowNumber,
                    0, 0);
                CellRangeAddress cellRangeAddress3 = new CellRangeAddress(rowNumber - dateRowCount, rowNumber,
                    2, 2);
                CellRangeAddress cellRangeAddress2 = new CellRangeAddress(rowNumber - storeRowCount, rowNumber,
                    1, 1);
                if (rowNumber - brandRowCount != rowNumber) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
                }
                if (rowNumber - dateRowCount != rowNumber) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
                }
                if (rowNumber - storeRowCount != rowNumber) {
                    writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
                }
                dateRowCount = 0;
                storeRowCount = 0;
                brandRowCount = 0;
            }
        }


        oldRow = row;


    }


    public void mergeBrand(WriteSheetHolder writeSheetHolder, int preRow) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(preRow - brandRowCount, preRow,
            0, 0);
        CellRangeAddress cellRangeAddress2 = new CellRangeAddress(preRow - storeRowCount, preRow,
            1, 1);
        CellRangeAddress cellRangeAddress3 = new CellRangeAddress(preRow - dateRowCount, preRow,
            2, 2);

        if (preRow - brandRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
        }
        writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress);
        if (preRow - dateRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
        }
        if (preRow - storeRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
        }
        dateRowCount = 0;
        storeRowCount = 1;
        brandRowCount = 1;
    }

    public void mergeStore(WriteSheetHolder writeSheetHolder, int preRow) {
        CellRangeAddress cellRangeAddress2 = new CellRangeAddress(preRow - storeRowCount, preRow,
            1, 1);
        CellRangeAddress cellRangeAddress3 = new CellRangeAddress(preRow - dateRowCount, preRow,
            2, 2);
        if (preRow - dateRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
        }
        if (preRow - storeRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress2);
        }
        dateRowCount = 0;
        storeRowCount = 1;
    }


    public void mergeDate(WriteSheetHolder writeSheetHolder, int preRow) {
        CellRangeAddress cellRangeAddress3 = new CellRangeAddress(preRow - dateRowCount, preRow,
            2, 2);
        if (preRow - dateRowCount != preRow) {
            writeSheetHolder.getSheet().addMergedRegionUnsafe(cellRangeAddress3);
        }
        dateRowCount = 0;
    }


}
