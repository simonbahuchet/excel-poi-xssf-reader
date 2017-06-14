package org.sba.excel.poi.reader.callback;

/**
 * Callback for notifying sheet processing
 *
 * @author Simon Bahuchet
 */
interface ExcelSheetCallback {

    /**
     * Callback for Worksheet start
     *
     * @param sheetNum
     * @param sheetName
     */
    void startSheet(int sheetNum, String sheetName) throws Exception

    /**
     * Callback for Worksheet end
     */
    void endSheet(int sheetNum, String sheetName) throws Exception

}