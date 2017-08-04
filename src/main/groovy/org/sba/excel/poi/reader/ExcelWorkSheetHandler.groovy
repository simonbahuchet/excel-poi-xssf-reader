package org.sba.excel.poi.reader

import groovy.util.logging.Slf4j
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
import org.apache.poi.xssf.usermodel.XSSFComment
import org.sba.excel.poi.reader.callback.ExcelRowContentCallback
import org.sba.excel.poi.reader.callback.ExcelRowSkippedCallback

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

    private static final int HEADER_ROW = 0

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
    ColumnMapping mappingMode = ColumnMapping.HEADER

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

        if (shouldSkipRow()) {
            return
        }

        if (considerHeaders()) {
            this.columnHeaders = new LinkedHashMap<String, String>()

        } else {
            this.currentRowMap = new LinkedHashMap<String, String>()

            // Add column header as key into current row map so that each entry
            // will exist. This ensures each column header will be in the "currentRowMap"
            // when passed to the user callback. Remember, the 'column headers map key' is the actual cell
            // column reference, it's value is the file column header value.
            // In the 'cell' method below, this empty string will be overwritten
            // with the file row value (if has one, else remains empty).
            for (String columnHeader : this.columnHeaders.values()) {
                this.currentRowMap.put(columnHeader, "")
            }
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
            String columnHeader = this.columnHeaders.get(columnReference)
            if (columnHeader) {
                this.currentRowMap.put(columnHeader, formattedValue)
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
        if (!shouldSkipRow() && rowNum <= HEADER_ROW) {
            //This is not yet the end, my friend
            return
        }

        // The row has NOT been skipped, call the appropriate callback
        if (!shouldSkipRow()) {
            try {
                //log.debug "Row has been read. rowNum=$rowNum, map=$currentRowMap. Let's call the rowCallback"
                this.rowCallback.processRow(rowNum, currentRowMap)
            } catch (Exception e) {
                throw new RuntimeException("Error invoking callback", e)
            }
            return
        }

        // The row has been skipped, call the appropriate callback
        if (this.rowSkippedCallback) {
            try {
                log.debug "Row $rowNum has been skipped, call the rowSkippedCallback"

                this.rowSkippedCallback.skip(rowNum)
            } catch (Exception e) {
                throw new RuntimeException("Error invoking callback", e)
            }
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
        MANUAL,
        HEADER
    }

    /**
     * @return true if we must consider the column values, of the current row, as headers
     */
    boolean considerHeaders() {
        boolean headerMode = this.mappingMode == ColumnMapping.HEADER && HEADER_ROW == this.currentRow
        if (headerMode && !this.columnHeaders) {
            throw new IllegalStateException("Cannot read any row without any mapping (ie. column names). Please " +
                    "switch to the HEADER mode or set a ColumnHeaders map")
        }
        return headerMode
    }

    /**
     * Set the mapping that will be applied for each row
     * @param columnHeaders
     */
    void setColumnHeaders(LinkedHashMap<String, String> columnHeaders) {
        mappingMode = ColumnMapping.MANUAL
        this.columnHeaders = columnHeaders
    }
}