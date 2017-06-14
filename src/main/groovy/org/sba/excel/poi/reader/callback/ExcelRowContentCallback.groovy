package org.sba.excel.poi.reader.callback

/**
 * Callback for processing a single row from excel file. Map keys are same as first row header
 * columns.
 *
 * @author Simon Bahuchet
 */
interface ExcelRowContentCallback {

    /**
     * Process the whole row at once
     *
     * @param rowNum index of the row
     * @param skipped true if the row has been skipped
     * @param map content of the row or null if it has been skipped
     * @throws Exception
     */
    void processRow(int rowNum, Map<String, String> map) throws Exception
}