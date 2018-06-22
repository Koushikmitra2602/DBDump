import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelReader {
    public static final String SAMPLE_XLSX_FILE_PATH = "./db_names.xlsx";

    public static Map<String, Object> getTableList() throws IOException, InvalidFormatException {

        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        Sheet sheetProps = workbook.getSheet("connectionProperties");
        Sheet sheetTable = workbook.getSheet("tableNames");

        DataFormatter dataFormatter = new DataFormatter();
        List<String> tableList = new ArrayList<>();
        sheetTable.forEach(row -> {
            row.forEach(cell -> {
            	tableList.add(dataFormatter.formatCellValue(cell));
            });
        });
        
        Map<String, String> propsMap = new HashedMap<>();
        sheetProps.forEach(row -> {
            	propsMap.put(dataFormatter.formatCellValue(row.getCell(0)), dataFormatter.formatCellValue(row.getCell(1)));
        });

        // Closing the workbook
        workbook.close();
        
        Map<String, Object> finalList = new HashedMap<>();
        finalList.put("PROP", propsMap);
        finalList.put("TABLES", tableList);
        
        return finalList;
    }
}