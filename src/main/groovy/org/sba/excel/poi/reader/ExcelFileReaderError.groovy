package org.sba.excel.poi.reader

class ExcelFileReaderError {

    /** Code identifying the error */
    String code

    /** Default message */
    String defaultMessage

    /** Optional - sheet where the error occurred */
    String sheetName

    /** Optional - row number where the error occurred */
    int rowNum

    /** Optional - column (1, 2, 3 or A, B, ..) where the error occurred */
    String columnReference
}
