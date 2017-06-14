package org.sba.excel.poi.reader.callback;

/**
 * Callback for notifying the sheet is being skipped
 *
 * @author Simon Bahuchet
 */
interface ExcelSheetSkippedCallback {

    /**
     * Callback for Worksheet skip
     *
     * @param sheetNum
     * @param sheetName
     */
    void skip(int sheetNum, String sheetName) throws Exception
}