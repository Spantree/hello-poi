import org.apache.poi.ss.util.CellReference
import spock.lang.Specification

class FileParserSpec extends Specification {
    def fileInputStreamMultipleSheetsXlsx
    def fileInputStreamMultipleSheetsXls

    def fileParser

    def expected
    def expectedMultipleSheets

    void setup() {
        fileInputStreamMultipleSheetsXlsx = getFile("/TestFileMultipleSheets.xlsx")
        fileInputStreamMultipleSheetsXls = getFile("/TestFileMultipleSheets.xls")

        fileParser = FileParser.builder().build()
        expected = [:]

        for (column in 0..4) {
            def k1 = createCellReferenceWithSheet("Sheet1", 0, column)
            expected[k1] = "column${column+1}"
            for(row in 1..7) {
                def k2 = createCellReferenceWithSheet("Sheet1", row, column)
                expected[k2] = column + 1
            }
        }

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

    InputStream getFile(String path) {
        ClassLoader.class.getResourceAsStream(path)
    }

    String createCellReferenceWithSheet(String sheetName, Integer rowNum, Integer columnNum) {
        new CellReference(sheetName, rowNum, columnNum, true, true).formatAsString()
    }

    def getParsedValue(Map parsedOutput, String sheetName, int row, int column) {
        def cellRef = createCellReferenceWithSheet(sheetName, row, column)
        return parsedOutput[cellRef]
    }

    def "Can successfully read xlsx file"() {
        given:
        def file = getFile("/TestFile.xlsx")
        def sheetName = "Sheet1"

        when: "processing a xlsx file"
        def parsedOutput = fileParser.parse(file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        row | col | expectedValue
        0   | 0   | "column1"
        0   | 1   | "column2"
        0   | 2   | "column3"
        0   | 3   | "column4"
        0   | 4   | "column5"

        1   | 0   | 1
        1   | 1   | 2
        1   | 2   | 3
        1   | 3   | 4
        1   | 4   | 5

        2   | 0   | 1
        2   | 1   | 2
        2   | 2   | 3
        2   | 3   | 4
        2   | 4   | 5

        3   | 0   | 1
        3   | 1   | 2
        3   | 2   | 3
        3   | 3   | 4
        3   | 4   | 5

        4   | 0   | 1
        4   | 1   | 2
        4   | 2   | 3
        4   | 3   | 4
        4   | 4   | 5

        5   | 0   | 1
        5   | 1   | 2
        5   | 2   | 3
        5   | 3   | 4
        5   | 4   | 5

        6   | 0   | 1
        6   | 1   | 2
        6   | 2   | 3
        6   | 3   | 4
        6   | 4   | 5
    }

    def "Can successfully read xls file"() {
        given:
        def file = getFile("/TestFileOlderVersion.xls")
        def sheetName = "Sheet1"

        when: "processing a xls file"
        def parsedOutput = fileParser.parse(file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        row | col | expectedValue
        0   | 0   | "column1"
        0   | 1   | "column2"
        0   | 2   | "column3"
        0   | 3   | "column4"
        0   | 4   | "column5"

        1   | 0   | 1
        1   | 1   | 2
        1   | 2   | 3
        1   | 3   | 4
        1   | 4   | 5

        2   | 0   | 1
        2   | 1   | 2
        2   | 2   | 3
        2   | 3   | 4
        2   | 4   | 5

        3   | 0   | 1
        3   | 1   | 2
        3   | 2   | 3
        3   | 3   | 4
        3   | 4   | 5

        4   | 0   | 1
        4   | 1   | 2
        4   | 2   | 3
        4   | 3   | 4
        4   | 4   | 5

        5   | 0   | 1
        5   | 1   | 2
        5   | 2   | 3
        5   | 3   | 4
        5   | 4   | 5

        6   | 0   | 1
        6   | 1   | 2
        6   | 2   | 3
        6   | 3   | 4
        6   | 4   | 5
    }

    def "Can successfully parse the output of a formula in a xlsx file"() {
        given:
        def file = getFile("/TestFileFormulas.xlsx")
        def sheetName = "Sheet1"

        when: "processing a file with formulas"
        def parsedOutput = fileParser.parse(file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        row | col | expectedValue
        0   | 0   | "Addition"
        0   | 1   | "Division"
        0   | 2   | "Neighbor Multiplication"

        1   | 0   | 8.0
        1   | 1   | 9.0
        1   | 2   | 72.0
    }

    def "Can successfully parse the output of a formula in a xls file"() {
        given:
        def file = getFile("/TestFileFormulas.xls")
        def sheetName = "Sheet1"

        when: "processing a file with formulas"
        def parsedOutput = fileParser.parse(file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col)

        where:
        row | col | expectedValue
        0   | 0   | "Addition"
        0   | 1   | "Division"
        0   | 2   | "Neighbor Multiplication"

        1   | 0   | 8.0
        1   | 1   | 9.0
        1   | 2   | 72.0
    }

    def "Can successfully handle an offset in a xlsx file"() {
        given:
        def file = getFile("/TestFileOffset.xlsx")
        def sheetName = "Sheet2"

        when: "processing cells in a specific row and sheet"
        def parsedOutput = fileParser.parseAtOffset("1:0", file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        row | col | expectedValue
        0   | 0   | "First"
        0   | 1   | "Second"
        0   | 2   | "Third"
    }

    def "Can successfully handle an offset in a xls file"() {
        given:
        def file = getFile("/TestFileOffset.xls")
        def sheetName = "Sheet2"

        when: "processing cells in a specific row and sheet"
        def parsedOutput = fileParser.parseAtOffset("1:0", file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        row | col | expectedValue
        0   | 0   | "First"
        0   | 1   | "Second"
        0   | 2   | "Third"
    }

    def "Can successfully read a xlsx file with multiple sheets"() {
        given:
        def file = getFile("/TestFileMultipleSheets.xlsx")

        when: "processing a xlsx file"
        def parsedOutput = fileParser.parse(file)

        then: "returns a hashmap with the processed input"
        getParsedValue(parsedOutput, sheetName, row, col) == expectedValue

        where:
        sheetName | row | col | expectedValue
        "Sheet1"  |  0   | 0   | "column1"
        "Sheet1"  |  0   | 1   | "column2"
        "Sheet1"  |  0   | 2   | "column3"

        "Sheet1"  |  1   | 0   | 1.0
        "Sheet1"  |  1   | 1   | 2.0
        "Sheet1"  |  1   | 2   | 3.0

        "Sheet1"  |  2   | 0   | 1.0
        "Sheet1"  |  2   | 1   | 2.0
        "Sheet1"  |  2   | 2   | 3.0

        "Sheet2"  |0   | 0   | "column1"
        "Sheet2"  |0   | 1   | "column2"
        "Sheet2"  |0   | 2   | "column3"

        "Sheet2"  |1   | 0   | 1.0
        "Sheet2"  |1   | 1   | 4.0
        "Sheet2"  |1   | 2   | 9.0

        "Sheet2"  |2   | 0   | 1.0
        "Sheet2"  |2   | 1   | 4.0
        "Sheet2"  |2   | 2   | 9.0
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