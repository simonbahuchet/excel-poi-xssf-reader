package org.sba.excel.poi.reader.callback;

/**
 * Callback to perform any task before and after reading the excel file
 *
 * @author Simon Bahuchet
 */
interface ExcelReadCallback {

    /**
     * Callback for Worksheet start
     *
     * @param sheetNum
     * @param sheetName
     */
    void beforeReading() throws Exception

    /**
     * Callback for Worksheet end
     */
    void afterReading() throws Exception

}