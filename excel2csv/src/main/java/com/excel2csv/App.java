package com.excel2csv;

import java.nio.file.Paths;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.DirectoryStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class App 
{
    public static void main( String[] args )
    {
        /*
         * Generate paths
         * Process one file at a time 
         */

        // long startTime = System.currentTimeMillis();

        String absolutePath = "Path to data folder";

        String folderPath = Paths.get(absolutePath, "input").toString();
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(folderPath))) {

            for(Path path : stream){
                String inPath = path.toAbsolutePath().toString();
                String outFileName = path.getFileName().toString().split("\\.")[0];
                outFileName += ".csv";
                String outPath = Paths.get(absolutePath, "output", outFileName).toString();
                processFile(inPath, outPath);
            }
        }
        catch (Exception e ){
            e.printStackTrace();
        }

        // long endTime = System.currentTimeMillis();
        // System.out.println("Elapsed Time in seconds: "+ (endTime-startTime)/1000);
    }

    public static void processFile(String inputPath, String outputPath){
        /*
         * Read input excel file and write to outout csv file
         */

        InputStream inp = null;
        FileOutputStream output = null;
        try {
            inp = new FileInputStream(inputPath);
            output = new FileOutputStream(outputPath);
            Workbook wb = WorkbookFactory.create(inp);

            for(int i=0;i<wb.getNumberOfSheets();i++) {
                // System.out.println(wb.getSheetAt(i).getSheetName());
                ExcelReading.echoAsCSV(wb.getSheetAt(i), output);
            }
        } 

        catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
        } 
        catch (IOException ex) {
            Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
        } 
        finally {
            try {
                inp.close();
                output.close();
            } 
            catch (IOException ex) {
                Logger.getLogger(ExcelReading.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
}
