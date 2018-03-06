import org.apache.poi.ss.util.CellReference
import spock.lang.Specification


class FileParserSpec extends Specification {
    def fileInputStreamXlsx
    def fileInputStreamXls
    def fileInputStreamFormulas
    def fileInputStreamOffset
    def fileParser
    def expected
    def expectedFormulas
    def expectedOffset

    void setup() {
        fileInputStreamXlsx = ClassLoader.class.getResourceAsStream("/TestFile.xlsx")
        fileInputStreamXls = ClassLoader.class.getResourceAsStream("/TestFileOlderVersion.xls")
        fileInputStreamFormulas = ClassLoader.class.getResourceAsStream("/TestFileFormulas.xlsx")
        fileInputStreamOffset = ClassLoader.class.getResourceAsStream("/TestFileOffset.xlsx")
        fileParser = FileParser.builder().build()
        expected = [
                (CellReference.newInstance(0, 0).formatAsString()): "column1",
                (CellReference.newInstance(0, 1).formatAsString()): "column2",
                (CellReference.newInstance(0, 2).formatAsString()): "column3",
                (CellReference.newInstance(0, 3).formatAsString()): "column4",
                (CellReference.newInstance(0, 4).formatAsString()): "column5",
                (CellReference.newInstance(1, 0).formatAsString()): 1,
                (CellReference.newInstance(1, 1).formatAsString()): 2,
                (CellReference.newInstance(1, 2).formatAsString()): 3,
                (CellReference.newInstance(1, 3).formatAsString()): 4,
                (CellReference.newInstance(1, 4).formatAsString()): 5,
                (CellReference.newInstance(2, 0).formatAsString()): 1,
                (CellReference.newInstance(2, 1).formatAsString()): 2,
                (CellReference.newInstance(2, 2).formatAsString()): 3,
                (CellReference.newInstance(2, 3).formatAsString()): 4,
                (CellReference.newInstance(2, 4).formatAsString()): 5,
                (CellReference.newInstance(3, 0).formatAsString()): 1,
                (CellReference.newInstance(3, 1).formatAsString()): 2,
                (CellReference.newInstance(3, 2).formatAsString()): 3,
                (CellReference.newInstance(3, 3).formatAsString()): 4,
                (CellReference.newInstance(3, 4).formatAsString()): 5,
                (CellReference.newInstance(4, 0).formatAsString()): 1,
                (CellReference.newInstance(4, 1).formatAsString()): 2,
                (CellReference.newInstance(4, 2).formatAsString()): 3,
                (CellReference.newInstance(4, 3).formatAsString()): 4,
                (CellReference.newInstance(4, 4).formatAsString()): 5,
                (CellReference.newInstance(5, 0).formatAsString()): 1,
                (CellReference.newInstance(5, 1).formatAsString()): 2,
                (CellReference.newInstance(5, 2).formatAsString()): 3,
                (CellReference.newInstance(5, 3).formatAsString()): 4,
                (CellReference.newInstance(5, 4).formatAsString()): 5,
                (CellReference.newInstance(6, 0).formatAsString()): 1,
                (CellReference.newInstance(6, 1).formatAsString()): 2,
                (CellReference.newInstance(6, 2).formatAsString()): 3,
                (CellReference.newInstance(6, 3).formatAsString()): 4,
                (CellReference.newInstance(6, 4).formatAsString()): 5,
                (CellReference.newInstance(7, 0).formatAsString()): 1,
                (CellReference.newInstance(7, 1).formatAsString()): 2,
                (CellReference.newInstance(7, 2).formatAsString()): 3,
                (CellReference.newInstance(7, 3).formatAsString()): 4,
                (CellReference.newInstance(7, 4).formatAsString()): 5,
        ]
        expectedFormulas = [
                (CellReference.newInstance(0, 0).formatAsString()): "Addition",
                (CellReference.newInstance(0, 1).formatAsString()): "Division",
                (CellReference.newInstance(0, 2).formatAsString()): "Neighbor Multiplication",
                (CellReference.newInstance(1, 0).formatAsString()): 8.0,
                (CellReference.newInstance(1, 1).formatAsString()): 9.0,
                (CellReference.newInstance(1, 2).formatAsString()): 72.0,
        ]
        expectedOffset = [
                (CellReference.newInstance(0, 0).formatAsString()): "First",
                (CellReference.newInstance(0, 1).formatAsString()): "Second",
                (CellReference.newInstance(0, 2).formatAsString()): "Third"
        ]
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

    def "Can successfully parse the output of a formula"() {
        when: "processing a file with formulas"
            def parsedOutput = fileParser.parse(fileInputStreamFormulas)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedFormulas
    }

    def "Can successfully handle an offset"() {
        when: "processing cells in a specific row and sheet"
            def parsedOutput = fileParser.parseAtOffset("1:0", fileInputStreamOffset)
        then: "returns a hashmap with the processed input"
            parsedOutput == expectedOffset
    }
}