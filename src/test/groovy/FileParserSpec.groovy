import spock.lang.Specification


class FileParserSpec extends Specification {
    def fileInputStream

    void setup() {
        fileInputStream = ClassLoader.class.getResourceAsStream("/TestFile.xlsx")
//        def fileInputStream = new FileInputStream("/resources/TestFile.xlsx")
    }

    def "Can successfully read xlsx file"() {
        when: "processing a xlsx file"
            println("IN THE TEST")
            println(fileInputStream)
            FileParser.parse(fileInputStream)
        then: "returns file input stream as a string"
    }
}