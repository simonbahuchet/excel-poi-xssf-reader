package org.sba.excel.poi.reader

import org.sba.excel.poi.reader.callback.DefaultExcelSheetCallback
import org.sba.excel.poi.reader.callback.ExcelSheetSkippedCallback
import groovy.util.logging.Slf4j
import org.apache.poi.openxml4j.exceptions.OpenXML4JException
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.util.IOUtils
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler
import org.apache.poi.xssf.model.StylesTable
import org.sba.excel.poi.reader.callback.ExcelReadCallback
import org.sba.excel.poi.reader.callback.ExcelSheetCallback
import org.xml.sax.ContentHandler
import org.xml.sax.InputSource
import org.xml.sax.SAXException
import org.xml.sax.XMLReader

import javax.xml.parsers.ParserConfigurationException
import javax.xml.parsers.SAXParserFactory

/**
 * Generic Excel File(XLSX) Reading using Apache POI
 *
 * <p>
 * Inspired by Jeevanandam M. <a href="https://github.com/jeevatkm/excelReader"
 * >https://github.com/jeevatkm/excelReader</a>
 * </p>
 *
 * @author <a href="mailto:jeeva@myjeeva.com">Jeevanandam M.</a>
 * @author Simon Bahuchet
 */
@Slf4j
class ExcelReader {

    private static final int READ_ALL = -1

    // Context used to share objects between reader, handler and callbacks
    ExecutionContext executionContext = new ExecutionContext()

    OPCPackage xlsxPackage
    ExcelWorkSheetHandler sheetContentsHandler

    // Handler called every time a sheet is opened or closed
    ExcelSheetCallback sheetCallback

    // Handler called before/after iterating over the sheets
    ExcelReadCallback readCallback

    // Handler that will notify sheets that are being skipped
    ExcelSheetSkippedCallback sheetSkippedCallback

    // List of row indices to skip
    List<String> sheetsToSkip = []

    /**
     * Constructor: Microsoft Excel File (XSLX) Reader
     *
     * @param pkg a {@link OPCPackage} object - The package to run XLSX
     * @param sheetContentsHandler a {@link SheetContentsHandler} object - WorkSheet contents handler
     */
    ExcelReader(OPCPackage pkg, ExcelWorkSheetHandler sheetContentsHandler) {
        this.xlsxPackage = pkg
        this.sheetContentsHandler = sheetContentsHandler
    }

    /**
     * Constructor: Microsoft Excel File (XSLX) Reader
     *
     * @param filePath a {@link String} object - The path of XLSX file
     * @param sheetContentsHandler a {@link SheetContentsHandler} object - WorkSheet contents handler
     */
    ExcelReader(String filePath, ExcelWorkSheetHandler sheetContentsHandler) throws Exception {
        this(getOPCPackage(filePath), sheetContentsHandler)
    }

    /**
     * Constructor: Microsoft Excel File (XSLX) Reader
     *
     * @param file a {@link File} object - The File object of XLSX file
     * @param sheetContentsHandler a {@link SheetContentsHandler} object - WorkSheet contents handler
     */
    ExcelReader(File file, ExcelWorkSheetHandler sheetContentsHandler) throws Exception {
        this(getOPCPackage(file), sheetContentsHandler)
    }

    /**
     * Constructor: Microsoft Excel File (XSLX) Reader
     *
     * @param InputStream a {@link InputStream} object - The InputStream object of XLSX file
     * @param sheetContentsHandler a {@link SheetContentsHandler} object - WorkSheet contents handler
     */
    ExcelReader(InputStream inputStream, ExcelWorkSheetHandler sheetContentsHandler) throws Exception {
        this(getOPCPackage(inputStream), sheetContentsHandler)
    }

    /**
     * Processing all the WorkSheet from XLSX Workbook.
     *
     * @throws Exception
     */
    void run(Map params) throws Exception {
        executionContext.add(params)
        read(READ_ALL)
    }

    /**
     * Processing of particular WorkSheet (zero based) from XLSX Workbook.
     *
     * @param sheetNumber a int object
     * @throws Exception
     */
    void run(int sheetNumber, Map params) throws Exception {
        executionContext.add(params)
        read(sheetNumber)
    }

    /**
     * Read one specific sheet / all the sheets stream and close it.
     *
     * Call the different handlers as well:
     * - sheetSkippedCallback if the sheet has to be skipped
     * - sheetCallback.startSheet before reading
     * - sheetCallback.endSheet after reading
     *
     * @param sheetNumber
     * @throws RuntimeException
     */
    private void read(int sheetNumber) throws RuntimeException {

        // Rest the execution context and pass it to all the ExecutionContextAware objects
        [sheetContentsHandler, sheetCallback, sheetSkippedCallback, readCallback].each {
            if (it && ExecutionContextAware.isAssignableFrom(it.class)) {
                ((ExecutionContextAware) it).setExecutionContext(this.executionContext)
            }
        }

        ReadOnlySharedStringsTable strings
        try {
            strings = new ReadOnlySharedStringsTable(this.xlsxPackage)
            XSSFReader xssfReader = new XSSFReader(this.xlsxPackage)
            StylesTable styles = xssfReader.getStylesTable()
            XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData()

            readCallback?.beforeReading()

            for (int sheetIndex = 0; worksheets.hasNext(); sheetIndex++) {

                InputStream stream = worksheets.next()

                try {
                    if (worksheets.getSheetName() in sheetsToSkip) {
                        sheetSkippedCallback?.skip(sheetIndex, worksheets.getSheetName())
                        continue
                    }

                    this.sheetCallback?.startSheet(sheetIndex, worksheets.getSheetName())

                    if ((READ_ALL == sheetNumber) || (sheetIndex == sheetNumber)) {
                        readSheet(styles, strings, stream)
                    }

                    this.sheetCallback?.endSheet(sheetIndex, worksheets.getSheetName())

                } finally {
                    IOUtils.closeQuietly(stream)
                }
            }

            readCallback?.afterReading()
        } catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException e) {
            log.error(e.getMessage(), e.getCause())
            throw new ExcelFileParseException("Couldn't read XLS file: ", e)
        }
    }

    /**
     * Parses the content of one sheet using the specified styles and shared-strings tables.
     *
     * @param styles a {@link StylesTable} object
     * @param sharedStringsTable a {@link ReadOnlySharedStringsTable} object
     * @param sheetInputStream a {@link InputStream} object
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    private void readSheet(StylesTable styles, ReadOnlySharedStringsTable sharedStringsTable,
                           InputStream sheetInputStream) throws IOException, ParserConfigurationException, SAXException {

        SAXParserFactory saxFactory = SAXParserFactory.newInstance()
        XMLReader sheetParser = saxFactory.newSAXParser().getXMLReader()

        ContentHandler handler = new XSSFSheetXMLHandler(styles, sharedStringsTable, sheetContentsHandler, true)

        sheetParser.setContentHandler(handler)
        sheetParser.parse(new InputSource(sheetInputStream))
    }

    /**
     * Open an OPCPackage from the file
     * @param filePath
     * @return
     */
    private static OPCPackage getOPCPackage(String filePath) throws Exception {
        if (!filePath) {
            throw new IllegalArgumentException("File path cannot be null")
        }
        return getOPCPackage(new File(filePath))
    }

    /**
     * Opens an OPCPackage for the file
     * @param file
     * @return OPCPackage
     */
    private static OPCPackage getOPCPackage(File file) throws Exception {
        if (null == file || !file.canRead()) {
            throw new IllegalArgumentException("File object is null or it misses the READ permission")
        }

        return OPCPackage.open(new FileInputStream(file))
    }

    /**
     * Opens an OPCPackage for the fileInputStream
     * @param FileInputStream
     * @return OPCPackage
     */
    private static OPCPackage getOPCPackage(InputStream inputStream) throws Exception {
        if (null == inputStream) {
            throw new IllegalArgumentException("inputStream object is null")
        }
        return OPCPackage.open(inputStream)
    }
}