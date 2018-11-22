package com.smartsheet.samples;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;

// Add Maven library "com.smartsheet:smartsheet-sdk-java:2.2.3" to access Smartsheet Java SDK
import com.smartsheet.api.Smartsheet;
import com.smartsheet.api.SmartsheetFactory;
import com.smartsheet.api.models.Cell;
import com.smartsheet.api.models.Column;
import com.smartsheet.api.models.Row;
import com.smartsheet.api.models.Sheet;


public class RWSheet {
	
    static {
    	
        // These lines enable logging to the console
        System.setProperty("Smartsheet.trace.parts", "RequestBodySummary, ResponseBodySummary");
        System.setProperty("Smartsheet.trace.pretty", "true");
    
    }

    // The API identifies columns by Id, but it's more convenient to refer to column names
    private static HashMap<String, Long> columnMap = new HashMap<String, Long>();   // Map from friendly column name to column Id

    public static void main(final String[] args) {

        try {
        	
            // Get API access token from properties file or environment
            Properties prop = new Properties();
            prop.load(RWSheet.class.getClassLoader().getResourceAsStream("rwsheet.properties"));

            String accessToken = prop.getProperty("accessToken");
            if (accessToken == null || accessToken.isEmpty())
                accessToken = System.getenv("SMARTSHEET_ACCESS_TOKEN");
            if (accessToken == null || accessToken.isEmpty())
                throw new Exception("Must set API access token in rwsheet.properties file");

            // Initialize client
            Smartsheet ss = SmartsheetFactory.createDefaultClient(accessToken);

            /*
             * Import the Excel file into a sheet.
             * NOTE: This only creates a sheet object with a reference to the Excel file.  The sheet will need to be loaded later.
             * API: https://smartsheet-platform.github.io/api-docs/?java#import
             */
            
            Sheet sheet = ss.sheetResources().importXlsx("Sample Sheet.xlsx", "sample", 0, 0);

            /*
             * Populate the sheet object with the contents of the Excel file
             * API: https://smartsheet-platform.github.io/api-docs/?java#get-sheet
             */
            
            sheet = ss.sheetResources().getSheet(sheet.getId(), null, null, null, null, null, null, null);
            System.out.println("Loaded " + sheet.getRows().size() + " rows from sheet: " + sheet.getName());

            // Build the column map for later reference
            for (Column column : sheet.getColumns()) {
                columnMap.put(column.getTitle(), column.getId());
                System.out.println("Mapping column with title=" + column.getTitle() + " to Id=" + column.getId());
        	}
            
            // Accumulate rows needing update here
            ArrayList<Row> rowsToUpdate = new ArrayList<Row>();

            for (Row row : sheet.getRows()) {
                Row rowToUpdate = evaluateRowAndBuildUpdates(row);
                if (rowToUpdate != null)
                    rowsToUpdate.add(rowToUpdate);
            }

            if (rowsToUpdate.isEmpty()) {
                System.out.println("No updates required");
            } else {
                // Finally, write all updated cells back to Smartsheet
                System.out.println("Writing " + rowsToUpdate.size() + " rows back to sheet id " + sheet.getId());
                ss.sheetResources().rowResources().updateRows(sheet.getId(), rowsToUpdate);
                System.out.println("Done");
            }
        } catch (Exception ex) {
            System.out.println("Exception : " + ex.getMessage());
            ex.printStackTrace();
        }
    }

    /*
     * TODO: Replace the body of this loop with your code
     * This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
     *
     * Return a new Row with updated cell values, else null to leave unchanged
     */
    private static Row evaluateRowAndBuildUpdates(Row sourceRow) {
        
    	System.out.println("Begin evaluateRowAndBuildUpdates() with sourceRow=" + sourceRow.getId());
    	
    	Row rowToUpdate = null;

        // Find cell we want to examine
        Cell statusCell = getCellByColumnName(sourceRow, "Status");
        System.out.println("Evaluating statusCell with displayValue=" + statusCell.getDisplayValue());
        
        if ("Complete".equals(statusCell.getDisplayValue())) {
        	
            Cell remainingCell = getCellByColumnName(sourceRow, "Remaining");
            System.out.println("Found remainingCell with displayValue=" + remainingCell.getDisplayValue());
            
            // Skip if "Remaining" is already zero
            if (! "0".equals(remainingCell.getDisplayValue())) {
                
            	System.out.println("Need to update row #" + sourceRow.getRowNumber());

                Cell cellToUpdate = new Cell();
                cellToUpdate.setColumnId(columnMap.get("Remaining"));
                cellToUpdate.setValue(0);
                System.out.println("Setting value=" + cellToUpdate.getValue() + " for remainingCell");
                
                // Convert the cell to be updated to a list to make it compatible with the following row update call
                List<Cell> cellsToUpdate = Arrays.asList(cellToUpdate);

                rowToUpdate = new Row();
                rowToUpdate.setId(sourceRow.getId());
                rowToUpdate.setCells(cellsToUpdate);
                System.out.println("Updating row #" + sourceRow.getRowNumber() + " with columnName=" + remainingCell.getDisplayValue() + " columnValue=" + cellToUpdate.getValue());
                
            }
        }
        
    	System.out.println("End evaluateRowAndBuildUpdates() with sourceRow=" + sourceRow.getId());
        
        return rowToUpdate;
    }

    // Helper function to find cell in a row
    static Cell getCellByColumnName(Row row, String columnName) {
        
    	Long colId = columnMap.get(columnName);

        return row.getCells().stream()
                .filter(cell -> colId.equals((Long) cell.getColumnId()))
                .findFirst()
                .orElse(null);

    }

}
