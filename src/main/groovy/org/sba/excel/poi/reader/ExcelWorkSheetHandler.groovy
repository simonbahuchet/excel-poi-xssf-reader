package org.sba.excel.poi.reader

import groovy.util.logging.Slf4j
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
import org.apache.poi.xssf.usermodel.XSSFComment
import org.sba.excel.poi.reader.callback.ExcelRowContentCallback
import org.sba.excel.poi.reader.callback.ExcelRowSkippedCallback

import static org.sba.excel.poi.reader.ExcelWorkSheetHandler.ColumnMapping.HEADER_ROW_ONLY
import static org.sba.excel.poi.reader.ExcelWorkSheetHandler.ColumnMapping.REFERENCE_AND_HEADER_ROW

/**
 * <p>
 * Excel Worksheet Handler for XML SAX parsing (.xlsx document model) <a
 * href="http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api"
 * >http://poi.apache.org/spreadsheet/how-to.html#xssf_sax_api</a>
 * </p>
 *
 * <p>
 * Inspired by Jeevanandam M. <a href="https://github.com/jeevatkm/excelReader"
 * >https://github.com/jeevatkm/excelReader</a>
 * </p>
 *
 * <p>
 * <strong>Usage:</strong> Provide a {@link ExcelRowContentCallback} callback that will be
 * provided a map representing a row of data from the file. The keys will be the column headers and values
 * the row data. Your callback class encapsulates any business logic for processing the string data
 * into dates, numbers, etc to allow full customization of the parsing and processing logic.
 * </p>
 *
 * @author https://github.com/DouglasCAyers
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam M.</a>
 * @author Simon Bahuchet
 */
@Slf4j
class ExcelWorkSheetHandler implements SheetContentsHandler, ExecutionContextAware {

    // mappingMode is HEADER by default => header row is the first row by default and can be changed by client
    int headerRow = 0

    // once an entire row of data has been read, pass map to this callback for processing
    private ExcelRowContentCallback rowCallback

    // map of column references => column headers (eg, 'A' => 'Product Title' )
    private int currentRow

    // LinkedHashMaps are used so iteration order is predictable over insertion order
    private LinkedHashMap<String, String> currentRowMap

    // List of row indices to skip
    List<Integer> rowsToSkip = []

    // If a row is being skipped, the following callback is called
    ExcelRowSkippedCallback rowSkippedCallback

    // Way to identify the keys for the row mapping -> are they identified from within a header row? or manually set
    ColumnMapping mappingMode = HEADER_ROW_ONLY

    // map of column headers => row values (eg, 'A' => 'White Shirts' )
    LinkedHashMap<String, String> columnHeaders

    // Context holding the objects shared between reader, handler and callbacks
    ExecutionContext executionContext

    /**
     * Default constructor
     * @param rowCallbackHandler
     */
    ExcelWorkSheetHandler(ExcelRowContentCallback rowCallbackHandler) {
        this.rowCallback = rowCallbackHandler
    }

    @Override
    void setExecutionContext(ExecutionContext executionContext) {
        this.executionContext = executionContext

        // Set the context of its "child objects"
        [rowSkippedCallback, rowCallback].each {
            if(it && ExecutionContextAware.isAssignableFrom(it.class)){
                ((ExecutionContextAware)it).setExecutionContext(this.executionContext)
            }
        }
    }

    /**
     * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler#startRow(int)
     */
    @Override
    void startRow(int rowNum) {

        this.currentRow = rowNum

        if (shouldSkipRow() && currentRow != headerRow) {
            return
        }

        if (considerHeaders()) {
            this.columnHeaders = new LinkedHashMap<String, String>()
        } else {
            this.currentRowMap = new LinkedHashMap<String, String>()
        }
    }

    /**
     * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler cell(java.lang.String,
     *      java.lang.String)
     */
    @Override
    void cell(String cellReference, String formattedValue, XSSFComment comment) {

        if (shouldSkipRow()) {
            return
        }

        // Note, POI will not invoke this method if the cell
        // is blank or if it detects there's no more data in the row.
        // So don't count on this being invoked the same number of times each
        // row. That's another reason why in above code we ensure each column header
        // is in the 'currentRowMap'.

        if (considerHeaders()) {
            this.columnHeaders.put(getColumnReference(cellReference), formattedValue)
        } else {
            String columnReference = getColumnReference(cellReference)
            String columnName = (mappingMode == REFERENCE_AND_HEADER_ROW) ? columnReference :  this.columnHeaders.get(columnReference)
            if (columnName) {
                this.currentRowMap.put(columnName, formattedValue)
            } else {
                //log.debug "Ignore cell [$cellReference:$formattedValue] because the column $columnReference is not mapped"
            }
        }
    }

    /**
     * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler endRow()
     */
    @Override
    void endRow(int rowNum) {
        // if is row to skip
        if (shouldSkipRow()) {
            // The row has been skipped, call the appropriate callback
            if (this.rowSkippedCallback) {
                try {
                    log.debug "Row $rowNum has been skipped, call the rowSkippedCallback"

                    this.rowSkippedCallback.skip(rowNum)
                } catch (Exception e) {
                    throw new RuntimeException("Error invoking callback", e)
                }
            }
            return
        }

        // if is header row
        if (mappingMode in [HEADER_ROW_ONLY, REFERENCE_AND_HEADER_ROW] && rowNum == headerRow) {
            // skip header row
            return
        }

        // otherwise, we have to process this row normally (call the appropriate callback)
        try {
            //log.debug "Row has been read. rowNum=$rowNum, map=$currentRowMap. Let's call the rowCallback"
            this.rowCallback.processRow(rowNum, currentRowMap)
        } catch (Exception e) {
            throw new RuntimeException("Error invoking callback", e)
        }

    }

    /**
     * @see org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler headerFooter(java.lang.String,
     *      boolean, java.lang.String)
     */
    @Override
    void headerFooter(String text, boolean isHeader, String tagName) {
        if (shouldSkipRow()) {
            return
        }

        // Not Used
    }

    /**
     * Returns the alphabetic column reference from this cell reference. Example: Given 'A12' returns
     * 'A' or given 'BA205' returns 'BA'
     */
    private String getColumnReference(String cellReference) {
        return cellReference ? cellReference.split('[0-9]*$')[0] : ""
    }

    /**
     * Should we skip the row / is its index member of the skipped rows list
     * @return
     */
    private boolean shouldSkipRow() {
        return currentRow in rowsToSkip
    }

    enum ColumnMapping {
        // CONFIGURED if we want to specify column header name manually
        CONFIGURED,
        // HEADER_ROW_ONLY if we only want to use column header name referenced by header row
        HEADER_ROW_ONLY,
        // REFERENCE_AND_HEADER_ROW if we want to manage excel column reference (A, B, C, ...) as column header
        // and also header row referenced by excel column reference
        REFERENCE_AND_HEADER_ROW
    }

    /**
     * @return true if we must consider the column values, of the current row, as headers
     */
    boolean considerHeaders() {
        return this.mappingMode in [HEADER_ROW_ONLY, REFERENCE_AND_HEADER_ROW] && headerRow == this.currentRow
    }

    /**
     * Set the mapping that will be applied for each row
     * @param columnHeaders
     */
    void setColumnHeaders(LinkedHashMap<String, String> columnHeaders) {
        mappingMode = ColumnMapping.CONFIGURED
        this.columnHeaders = columnHeaders
    }
}