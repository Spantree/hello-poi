import org.apache.poi.ss.util.CellReference
import spock.lang.Specification

class FileParserSpec extends Specification {
    def fileInputStreamXlsx
    def fileInputStreamXls
    def fileInputStreamFormulasXlsx
    def fileInputStreamOffsetXlsx
    def fileInputStreamFormulasXls
    def fileInputStreamOffsetXls
    def fileInputStreamMultipleSheetsXlsx
    def fileInputStreamMultipleSheetsXls

    def fileParser

    def expected
    def expectedFormulas
    def expectedOffset
    def expectedMultipleSheets

    void setup() {
        fileInputStreamXlsx = ClassLoader.class.getResourceAsStream("/TestFile.xlsx")
        fileInputStreamXls = ClassLoader.class.getResourceAsStream("/TestFileOlderVersion.xls")
        fileInputStreamFormulasXlsx = ClassLoader.class.getResourceAsStream("/TestFileFormulas.xlsx")
        fileInputStreamOffsetXlsx = ClassLoader.class.getResourceAsStream("/TestFileOffset.xlsx")
        fileInputStreamFormulasXls = ClassLoader.class.getResourceAsStream("/TestFileFormulas.xls")
        fileInputStreamOffsetXls = ClassLoader.class.getResourceAsStream("/TestFileOffset.xls")
        fileInputStreamMultipleSheetsXlsx = ClassLoader.class.getResourceAsStream("/TestFileMultipleSheets.xlsx")
        fileInputStreamMultipleSheetsXls = ClassLoader.class.getResourceAsStream("/TestFileMultipleSheets.xls")

        fileParser = FileParser.builder().build()
        expected = [
                (createCellReferenceWithSheet("Sheet1", 0, 0)): "column1",
                (createCellReferenceWithSheet("Sheet1", 0, 1)): "column2",
                (createCellReferenceWithSheet("Sheet1", 0, 2)): "column3",
                (createCellReferenceWithSheet("Sheet1", 0, 3)): "column4",
                (createCellReferenceWithSheet("Sheet1", 0, 4)): "column5",
                (createCellReferenceWithSheet("Sheet1", 1, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 1, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 1, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 1, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 1, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 2, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 2, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 2, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 2, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 2, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 3, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 3, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 3, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 3, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 3, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 4, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 4, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 4, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 4, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 4, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 5, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 5, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 5, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 5, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 5, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 6, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 6, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 6, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 6, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 6, 4)): 5,
                (createCellReferenceWithSheet("Sheet1", 7, 0)): 1,
                (createCellReferenceWithSheet("Sheet1", 7, 1)): 2,
                (createCellReferenceWithSheet("Sheet1", 7, 2)): 3,
                (createCellReferenceWithSheet("Sheet1", 7, 3)): 4,
                (createCellReferenceWithSheet("Sheet1", 7, 4)): 5,
        ]
        expectedFormulas = [
                (createCellReferenceWithSheet("Sheet1", 0, 0)): "Addition",
                (createCellReferenceWithSheet("Sheet1", 0, 1)): "Division",
                (createCellReferenceWithSheet("Sheet1", 0, 2)): "Neighbor Multiplication",
                (createCellReferenceWithSheet("Sheet1", 1, 0)): 8.0,
                (createCellReferenceWithSheet("Sheet1", 1, 1)): 9.0,
                (createCellReferenceWithSheet("Sheet1", 1, 2)): 72.0,
        ]
        expectedOffset = [
                (createCellReferenceWithSheet("Sheet2", 0, 0)): "First",
                (createCellReferenceWithSheet("Sheet2", 0, 1)): "Second",
                (createCellReferenceWithSheet("Sheet2", 0, 2)): "Third"
        ]

        expectedMultipleSheets = [
                (createCellReferenceWithSheet("Sheet1", 0,0)): "column1",
                (createCellReferenceWithSheet("Sheet1", 0,1)): "column2",
                (createCellReferenceWithSheet("Sheet1", 0,2)): "column3",
                (createCellReferenceWithSheet("Sheet1", 1,0)): 1.0,
                (createCellReferenceWithSheet("Sheet1", 1,1)): 2.0,
                (createCellReferenceWithSheet("Sheet1", 1, 2)): 3.0,
                (createCellReferenceWithSheet("Sheet1", 2, 0)): 1.0,
                (createCellReferenceWithSheet("Sheet1", 2, 1)): 2.0,
                (createCellReferenceWithSheet("Sheet1", 2, 2)): 3.0,
                (createCellReferenceWithSheet("Sheet2", 0, 0)): "column1",
                (createCellReferenceWithSheet("Sheet2", 0, 1)): "column2",
                (createCellReferenceWithSheet("Sheet2", 0, 2)): "column3",
                (createCellReferenceWithSheet("Sheet2", 1, 0)): 1.0,
                (createCellReferenceWithSheet("Sheet2", 1, 1)): 4.0,
                (createCellReferenceWithSheet("Sheet2", 1, 2)): 9.0,
                (createCellReferenceWithSheet("Sheet2", 2, 0)): 1.0,
                (createCellReferenceWithSheet("Sheet2", 2, 1)): 4.0,
                (createCellReferenceWithSheet("Sheet2", 2, 2)): 9.0,
                (createCellReferenceWithSheet("Sheet3", 0, 0)): "column1",
                (createCellReferenceWithSheet("Sheet3", 0, 1)): "column2",
                (createCellReferenceWithSheet("Sheet3", 0, 2)): "column3",
                (createCellReferenceWithSheet("Sheet3", 1, 0)): 1.0,
                (createCellReferenceWithSheet("Sheet3", 1, 1)): 8.0,
                (createCellReferenceWithSheet("Sheet3", 1, 2)): 27.0,
                (createCellReferenceWithSheet("Sheet3", 2, 0)): 1.0,
                (createCellReferenceWithSheet("Sheet3", 2, 1)): 8.0,
                (createCellReferenceWithSheet("Sheet3", 2, 2)): 27.0
        ]
    }

    private String createCellReferenceWithSheet(String sheetName, Integer rowNum, Integer columnNum) {
        new CellReference(sheetName, rowNum, columnNum, true, true).formatAsString()
    }


    def "Can successfully read xlsx file"() {
        when: "processing a xlsx file"
            def parsedOutput = fileParser.parse(fileInputStreamXlsx)
        then: "returns a hashmap with the processed input"
            parsedOutput == expected
    }

    def "Can successfully read xls file"() {
        when: "processing a xls file"
            def parsedOutput = fileParser.parse(fileInputStreamXls)
        then: "returns a hashmap with the processed input"
            parsedOutput == expected
    }

    def "Can successfully parse the output of a formula in a xlsx file"() {
        when: "processing a file with formulas"
            def parsedOutput = fileParser.parse(fileInputStreamFormulasXlsx)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedFormulas
    }

    def "Can successfully handle an offset in a xlsx file"() {
        when: "processing cells in a specific row and sheet"
            def parsedOutput = fileParser.parseAtOffset("1:0", fileInputStreamOffsetXlsx)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedOffset
    }

    def "Can successfully parse the output of a formula in a xls file"() {
        when: "processing a file with formulas"
            def parsedOutput = fileParser.parse(fileInputStreamFormulasXls)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedFormulas
    }

    def "Can successfully handle an offset in a xls file"() {
        when: "processing cells in a specific row and sheet"
            def parsedOutput = fileParser.parseAtOffset("1:0", fileInputStreamOffsetXls)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedOffset
    }

    def "Can successfully read a xlsx file with multiple sheets"() {
        when: "processing a xlsx file"
            def parsedOutput = fileParser.parse(fileInputStreamMultipleSheetsXlsx)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedMultipleSheets
    }

    def "Can successfully read a xls file with multiple sheets"() {
        when: "processing a xls file"
            def parsedOutput = fileParser.parse(fileInputStreamMultipleSheetsXls)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedMultipleSheets
    }

//    TODO: Write a test for a formula which refers to multiple cells, such as summing up a column
    // TODO: Write a test for multiple sheet workbooks where the sheets have been renamed
}