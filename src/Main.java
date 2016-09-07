import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.util.Queue;
import java.util.Stack;

/**
 * Created by Xavier on 9/2/16.
 */
public class Main {

    XSSFWorkbook activeSync;
    XSSFSheet asSheet;

    XSSFWorkbook infin;
    XSSFSheet infinSheet;

    String name;
    String secondName;
    Stack<String> names = new Stack();
    Stack<XSSFRow> leftover = new Stack();



    public static void main(String[] args) throws IOException, InvalidFormatException {
        Main spreadSheet = new Main();
        XSSFWorkbook output = spreadSheet.read();
        FileOutputStream out = new FileOutputStream("/Users/Xavier/Documents/Work Stuff/MasterList.xlsx");
        output.write(out);


    }

    public XSSFWorkbook read() throws IOException, InvalidFormatException {

        File infinFile = new File("/Users/Xavier/Documents/Work Stuff/InfiniumBackup.xlsx");
        File actFile = new File("/Users/Xavier/Documents/Work Stuff/ActiveSyncBackup.xlsx");

        infin = new XSSFWorkbook(infinFile);
        activeSync = new XSSFWorkbook(actFile);
        asSheet = activeSync.getSheetAt(0);
        infinSheet = infin.getSheetAt(0);
        int count = infinSheet.getPhysicalNumberOfRows();
        int originalCount = count; // reference to past value to start "leftover" function loop;


        Cell cell;
        for(int i = 0; i < asSheet.getPhysicalNumberOfRows(); i++){
            cell = asSheet.getRow(i).getCell(0);
            name = getName(cell); // We now have a name to compare to spreadsheet 2;
            names.push(name);
        }
        String prevName = " "; // keeps track of previous name to get rid of doubles.

        while(!names.empty()) {
            name = names.peek();
            if(name.equalsIgnoreCase(prevName)){
                prevName = names.peek();
                name = name + " " + Math.random() * 10;
            }
            for (int j = 0; j <= infinSheet.getPhysicalNumberOfRows(); j++) { //starts from end because stack
                if (j == infinSheet.getPhysicalNumberOfRows()) { //reach end of list, not found
                    System.out.println("Row Added to stack:");
                    //createNewRow(asSheet, names.size() - 1, infinSheet, 1);
                    leftover.push(getRow(asSheet, names.size() -1));
                    count++;
                    break;
                } else {
                    if (infinSheet.getRow(j) != null) {
                        secondName = infinSheet.getRow(j).getCell(0).toString() + " ";
                        secondName += infinSheet.getRow(j).getCell(1).toString();
                        //System.out.println(name + " --- " + secondName);
                        if (secondName.equalsIgnoreCase(name)) {
                            //System.out.println("true");
                            copyToNewSpreadsheet(asSheet, names.size()-1, infinSheet, j);
                            break;
                        }
                    }
                }
            }
            prevName = names.pop();
        }
            finishItOff(leftover, infinSheet,originalCount);
            return infin;

    }
    /*
    creates new row and inserts it at the beginning of the new spreadsheet
     */
    public void createNewRow(XSSFSheet worksheet, int sourceRowNum, XSSFSheet destination, int destinationRowNum) throws IOException {
       try {
           XSSFRow newRow = destination.getRow(destinationRowNum);

           XSSFRow sourceRow = worksheet.getRow(sourceRowNum);

           // If the row exist in destination, push down all rows by 1 else create a new row
           if (newRow != null) {
               destination.shiftRows(1, destination.getLastRowNum(), 1);
           } else {
               newRow = worksheet.createRow(destinationRowNum);
           }

            int column = 16;
           // Loop through source columns to add to new row
           for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
               // Grab a copy of the old/new cell
               XSSFCell oldCell = sourceRow.getCell(i);
               XSSFCell newCell = newRow.createCell(i + column, Cell.CELL_TYPE_STRING);

               // If the old cell is null jump to next cell
               if (oldCell == null) {
                   newCell = null;
                   continue;
               }

               // If there is a cell comment, copy
               if (oldCell.getCellComment() != null) {
                   newCell.setCellComment(oldCell.getCellComment());
               }

               // If there is a cell hyperlink, copy
               if (oldCell.getHyperlink() != null) {
                   newCell.setHyperlink(oldCell.getHyperlink());
               }
               newCell.setCellValue(oldCell.toString());

           }
       }
        catch(Exception e){
            e.printStackTrace();
        }

    }

    public XSSFRow getRow(XSSFSheet sheet, int rowNumber){
        return sheet.getRow(rowNumber);
    }

    public void finishItOff(Stack<XSSFRow> rows, XSSFSheet sheet, int rowNumber){

        /*for(int i = beginLoop; i <endLoop; i++){
            XSSFRow newRow = createRowAtEnd(sheet,"Name Added",i);
            XSSFRow currentRow = rows.pop();
            newRow.copyRowFrom(currentRow, new CellCopyPolicy());
        }*/
        String name = "";
        XSSFCell first;
        XSSFCell last;
        while(!rows.empty()){
            name = rows.peek().getCell(0).toString();
            String[] full = name.split(",");
            XSSFRow newRow = createRowAtEnd(sheet, name, rowNumber);
            XSSFRow oldRow = rows.pop();
            last = newRow.createCell(0, Cell.CELL_TYPE_STRING);
            name = "";
            for(String x : full){
                name+= x;
            }
            last.setCellValue(name);




            int column = 16;
            for(int i = 0; i < oldRow.getLastCellNum(); i++){
                XSSFCell oldCell = oldRow.getCell(i);
                XSSFCell newCell = newRow.createCell(i + column, Cell.CELL_TYPE_STRING);
                if(oldCell == null){
                    newCell = null;
                    continue;
                }
                newCell.setCellValue(oldCell.toString());
            }
            rowNumber++;
        }
    }

    public XSSFRow createRowAtEnd(XSSFSheet sheet, String name, int rowNumber){
        XSSFRow newRow = sheet.createRow(rowNumber);
        String[] fullName = name.split(" ");
        XSSFCell last = newRow.createCell(0,Cell.CELL_TYPE_STRING);
        XSSFCell first = newRow.createCell(1, Cell.CELL_TYPE_STRING);
        last.setCellValue(fullName[0]);
        first.setCellValue(fullName[1]);
        return newRow;
    }

    public void copyToNewSpreadsheet(XSSFSheet asSheet, int originalRow, XSSFSheet infinSheet, int destRow) throws IOException {
        try {
            /*XSSFRow toBeCopied = original.getRow(originalRow);
            XSSFRow destinationRow = destination.getRow(destRow);
            int column = 16; // when the activeSync should start being copied
            if (toBeCopied != null) {
                for (int i = 0; i < toBeCopied.getLastCellNum(); i++) {
                    Cell oldCell = toBeCopied.getCell(i);
                    Cell newCell = destinationRow.createCell(column + i);

                    if (oldCell == null) {
                        newCell = null;
                        continue;
                    }

                    newCell.setCellValue(oldCell.toString());
                }
            }*/


            XSSFRow newRow = infinSheet.getRow(destRow);
            XSSFRow sourceRow = asSheet.getRow(originalRow);


            int column = 16;
            // Loop through source columns to add to new row
            for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                // Grab a copy of the old/new cell
                XSSFCell oldCell = sourceRow.getCell(i);
                XSSFCell newCell = newRow.createCell(i + column, Cell.CELL_TYPE_STRING);

                // If the old cell is null jump to next cell
                if (oldCell == null) {
                    newCell = null;
                    continue;
                }
                newCell.setCellValue(oldCell.toString());
            }
        }
        catch(Exception e ){
            e.printStackTrace();
        }
    }

    public String getName(Cell cell){
        String[] fullName = cell.toString().split(",");
        String name = "";

        for(String x : fullName){
            name += x;
        }

        return name.trim();
    }

   /* public void getLeftovers(Stack<String> leftovers, XSSFSheet sheet, XSSFSheet original, int originalCount, int count){
        while(!leftovers.empty()){
            for(int i = originalCount; i <= count; i++){
                String name = sheet.getRow(i).getCell(0).toString() + " ";
                name+= sheet.getRow(i).getCell(1).toString();
                if(leftovers.peek().equalsIgnoreCase(name)){
                    copyToNewSpreadsheet(sheet,i);
                }
            }
        }
    }*/
}
