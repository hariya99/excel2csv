package com.excel2csv;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReading {

    public static void echoAsCSV(Sheet sheet, FileOutputStream output) {
        Row row = null;
        StringBuilder outRow = new StringBuilder();
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            // outRow = "";
            outRow.delete( 0, outRow.length() );
            for (int j = 0; j < row.getLastCellNum(); j++) {
                // System.out.print(row.getCell(j) + "|");
                outRow.append(row.getCell(j) + "|");
            }
            try{
                outRow.append("\n");
                output.write(outRow.toString().getBytes());
            }
            catch(IOException e){
                e.getStackTrace();
            }
            
            // System.out.println();
        }
    }
}
