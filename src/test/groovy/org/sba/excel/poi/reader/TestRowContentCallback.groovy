package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.ExcelRowContentCallback
import groovy.util.logging.Slf4j

@Slf4j
class TestRowContentCallback implements ExcelRowContentCallback {

    int numberOfCalls = 0

    Map<Integer, Map<String, String>> rowMaps = [:]

    @Override
    void processRow(int rowIndex, Map<String, String> map) throws Exception {
        log.debug "n°$rowIndex => $map"
        numberOfCalls++
        rowMaps << [(rowIndex): map]
    }
}
