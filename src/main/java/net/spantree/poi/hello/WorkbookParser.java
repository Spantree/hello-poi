package net.spantree.poi.hello;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Collections;
import java.util.Map;
import java.util.stream.Stream;

public class WorkbookParser {

    public Stream<Map<String, Object>> parse(Workbook workbook) {
        return Streams.stream(workbook).flatMap(Streams::stream).map(this::parseRow);
    }

    private Map<String, Object> parseRow(Row row) {
        return Collections.emptyMap();
    }

}
