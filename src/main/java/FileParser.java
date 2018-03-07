import lombok.Builder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

@Builder
public class FileParser {

    public Map<String, Object> parseRow(String sheetName, Row row) throws IOException, InvalidFormatException {
        LinkedHashMap<String, Object> output = new LinkedHashMap<>();
            for (Cell cell : row) {
                CellReference cellReference = new CellReference(sheetName, row.getRowNum(), cell.getColumnIndex(), true, true);
                CellType cellType = cell.getCellTypeEnum();
                if (cell.getCellTypeEnum().equals(CellType.FORMULA)) {
                    cellType = cell.getCachedFormulaResultTypeEnum();
                }
                switch (cellType) {
                    case STRING:
                        output.put(cellReference.formatAsString(), cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        output.put(cellReference.formatAsString(), cell.getNumericCellValue());
                        break;
                }
            }
        return output;
    }

    public Map<String, Object> parseAtOffset(String offset, InputStream fileInputStream) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(fileInputStream);

//      Offset format will have the sheet number as the first item and the row number as the second item.  A colon will separate the two.  For example, 3:2 is sheet 3 and row 2.
        String[] splitOffset = offset.split(":");
        int sheetNum = Integer.parseInt(splitOffset[0]);
        int rowNum = Integer.parseInt(splitOffset[1]);
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Row row = sheet.getRow(rowNum);
        return parseRow(sheet.getSheetName(), row);
    }

    public Map<String, Object> parse(InputStream fileInputStream) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        LinkedHashMap<String, Object> output = new LinkedHashMap<>();
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                output.putAll(parseRow(sheet.getSheetName(), row));
            }
        }
        return output;
    }
}
