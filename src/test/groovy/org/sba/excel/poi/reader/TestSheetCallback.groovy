package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.ExcelSheetCallback
import groovy.util.logging.Slf4j

@Slf4j
class TestSheetCallback implements ExcelSheetCallback {

    int numberOfStarts = 0
    int numberOfEnds = 0

    @Override
    void startSheet(int sheetNum, String sheetName) {
        log.debug "Start sheet $sheetName"
        numberOfStarts++
    }

    @Override
    void endSheet(int sheetNum, String sheetName) {
        log.debug "End sheet $sheetName"
        numberOfEnds++
    }
}
