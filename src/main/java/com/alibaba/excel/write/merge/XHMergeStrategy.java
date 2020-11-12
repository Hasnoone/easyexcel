package com.alibaba.excel.write.merge;

import com.alibaba.excel.write.XHDemoData;
import com.alibaba.excel.write.handler.AbstractRowWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import org.apache.commons.collections4.MapUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * The regions of the loop merge
 *
 * @author Jiaju Zhuang
 */
public class XHMergeStrategy extends AbstractRowWriteHandler {
    /**
     * Each row
     */
    private List<XHDemoData> list;

    public XHMergeStrategy(List<XHDemoData> list) {
        this.list = list;
    }


    public Map<String, Map<String, Map<String, List<XHDemoData>>>> group() {
        Map<String, Map<String, Map<String, List<XHDemoData>>>> collect = list.stream().collect(Collectors.groupingBy(XHDemoData::getBrandName, Collectors.groupingBy(XHDemoData::getStoreName, Collectors.groupingBy(XHDemoData::getStoreName))));
        return collect;
    }




    @Override
    public void afterRowDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row,
        Integer relativeRowIndex, Boolean isHead) {
        if (isHead) {
            return;
        }
        Map<String, Map<String, Map<String, List<XHDemoData>>>> brandGroup = group();

        if (MapUtils.isNotEmpty(brandGroup)) {
            for (Map.Entry<String, Map<String, Map<String, List<XHDemoData>>>> brandItem : brandGroup.entrySet()) {
                String brandName = brandItem.getKey();
                Map<String, Map<String, List<XHDemoData>>> storeGroup = brandItem.getValue();
                if (MapUtils.isNotEmpty(storeGroup)) {
                    for (Map.Entry<String, Map<String, List<XHDemoData>>> storeItem : storeGroup.entrySet()) {
                        String storeName = storeItem.getKey();
                        Map<String, List<XHDemoData>> businessDateGroup = storeItem.getValue();
                        if (MapUtils.isNotEmpty(businessDateGroup)) {
                            for (Map.Entry<String, List<XHDemoData>> businessDateItem : businessDateGroup.entrySet()) {
                                String businessDate = businessDateItem.getKey();
                                List<XHDemoData> value = businessDateItem.getValue();
                                int size = value.size();



                            }
                        }
                    }
                }

            }
        }

    }

}
