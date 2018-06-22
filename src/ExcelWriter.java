import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelWriter {

    public static void writeExcel(ResultSet rs, Workbook workbook) throws IOException, InvalidFormatException, SQLException {
		List<String> columns = new ArrayList<>();
        ResultSetMetaData metaData = rs.getMetaData();
        for(int i = 1; i <= metaData.getColumnCount(); i ++){
        	columns.add(metaData.getColumnName(i));
        }
        
        String tableName = metaData.getTableName(1);
    	// Create a Workbook
        

        // Create a Sheet
        Sheet sheet = workbook.createSheet(tableName);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns.get(i));
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with employees data
        int rowNum = 1;
        while (rs.next()) {
        	Row row = sheet.createRow(rowNum++);
        	for(int i = 1; i <= metaData.getColumnCount(); i ++){
        		row.createCell(i-1)
        		.setCellValue(rs.getString(i));
        	}
        }

		// Resize all columns to fit the content size
        for(int i = 0; i < columns.size(); i++) {
            sheet.autoSizeColumn(i);
        }
    }
}