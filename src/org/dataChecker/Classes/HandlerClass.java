package org.dataChecker.Classes;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.*;

public class HandlerClass {
    
    Vector<Double> mainVector = new Vector<Double>();
    Vector<Double> fileVector = new Vector<Double>();
  
    public int readFile(File file,Vector<Double> vec){
        if(file.exists() && file.isFile());
        else
            return -1;
        try{
            FileInputStream in = new FileInputStream(file);
            //XSSF object for parsing and managing Excel file data
            XSSFWorkbook workbook = new XSSFWorkbook(in);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            
            while(rowIterator.hasNext()){
                XSSFRow row = (XSSFRow)rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                
                if(cellIterator.hasNext()){
                    XSSFCell cell = (XSSFCell)cellIterator.next();
                    //The primary key values are stored in first cell. read those values and put them in the map
                    double val = cell.getNumericCellValue();
                    vec.add(val);
                    
                    while(cellIterator.hasNext()){
                        cell = (XSSFCell)cellIterator.next();
                        //Check the cell type and format accordingly
                        switch (cell.getCellType()) 
                        {
                            case Cell.CELL_TYPE_NUMERIC:
                                System.out.print(cell.getNumericCellValue() + "\t");
                                break;
                            case Cell.CELL_TYPE_STRING:
                                System.out.print(cell.getStringCellValue() + "\t");
                                break;
                        }
                    }
                    System.out.println();
                    
                }
            }
            
            //Close the stream
            in.close();
        }
        catch(Exception e){
            e.printStackTrace();
        }
        
        return 0;
    }
    
    public int compareFiles(File mainFile, File secFile){
        //Read main file and make the vector of available values.
        int retMain = readFile(mainFile,mainVector);
        //Read the "file" file and make the vector of available values.
        int retFile = readFile(secFile,fileVector);
        
        if(retMain==-1 || retFile==-1){//File couldn't be read. Exit
            return -1;
        }
        
        
        Iterator<Double> fileVecIt = fileVector.iterator();
        while(fileVecIt.hasNext()){
            double no = fileVecIt.next();
            
            //Remove this no from mainVector such that mainVector 
            //contains only those values that are not in fileVector
            mainVector.removeElement(no);
        }
        
        //Write the remaining values to a new file to be stored on Desktop.
        writeFile(mainFile);
        
        return 0;
    }
    
    public int writeFile(File mainFile){    //Second input parameter is mainVector
        //Creating an instance for the workbook to write to
        XSSFWorkbook writeWorkbook = new XSSFWorkbook();
        XSSFSheet writeSheet = writeWorkbook.createSheet("Sheet 1");
        
                
        if(mainFile.exists() && mainFile.isFile());
        else
            return -1;
        try{
            FileInputStream in = new FileInputStream(mainFile);
            //XSSF object for parsing and managing Excel file data
            XSSFWorkbook mainWorkbook = new XSSFWorkbook(in);
            XSSFSheet mainSheet = mainWorkbook.getSheetAt(0);
            Iterator<Row> rowIterator = mainSheet.iterator();
            int writeRowNum=0;
            
            while(rowIterator.hasNext()){
                XSSFRow mainRow = (XSSFRow)rowIterator.next();
                Iterator<Cell> cellIterator = mainRow.cellIterator();
                
                XSSFRow writeRow = writeSheet.createRow(writeRowNum);
                int writeCellNum=0;
                
                if(cellIterator.hasNext()){
                    XSSFCell mainCell = (XSSFCell)cellIterator.next();  //Getting the cell at 0th position i.e. that contains the primary key
                    
                    if(mainVector.contains(mainCell.getNumericCellValue())){
                        XSSFCell writeCell = writeRow.createCell(writeCellNum);
                        writeCellNum++;
                        writeCell.setCellValue(mainCell.getNumericCellValue());
                        
                        while(cellIterator.hasNext()){
                            mainCell = (XSSFCell)cellIterator.next();
                            writeCell = writeRow.createCell(writeCellNum);
                            writeCellNum++;
                            //Check the cell type and format accordingly
                            switch (mainCell.getCellType()) 
                            {
                                case Cell.CELL_TYPE_NUMERIC:
                                    writeCell.setCellValue(mainCell.getNumericCellValue());
                                    System.out.print(mainCell.getNumericCellValue() + "\t");
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    writeCell.setCellValue(mainCell.getStringCellValue());
                                    System.out.print(mainCell.getStringCellValue() + "\t");
                                    break;
                            }
                        }
                        
                        writeRowNum++; //Increment value of row only if there is a value to be added.
                    }
                    
                }
                System.out.println();
            }
            
            //Close the stream
            in.close();
        }
        catch(Exception e){
            e.printStackTrace();
        }
        
        try{
            File writeFile = new File(System.getProperty("user.home"), "Documents/IEEEDataCheckerFile.xlsx");
            FileOutputStream out = new FileOutputStream(writeFile);
            writeWorkbook.write(out);
            out.close(); 
            System.out.println("Data has been successfully written and file has been created");
        }
        catch(Exception e){
            e.printStackTrace();
        }
        
        return -1;
    }
    
}
