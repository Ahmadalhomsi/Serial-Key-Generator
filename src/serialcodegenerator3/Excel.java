/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package serialcodegenerator3;

import java.util.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import javax.swing.JOptionPane;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author pooqw
 */
public class Excel {

    ArrayList<Key> keys = new ArrayList<Key>();
    int cCount = -1;

    public ArrayList<Key> ExcelReader(String path, boolean Unlimited) throws IOException, Exception {

        ArrayList<Key> ReadingList = new ArrayList<>();
        ArrayList<String> ColumnsNames = new ArrayList<>();

        try {

            // Reading file from local directory
            FileInputStream file = new FileInputStream(
                    new File(path));

            // Create Workbook instance holding reference to
            // .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            // Till there is an element condition holds true
            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                // For each row, iterate through all the
                // columns
                Iterator<Cell> cellIterator
                        = row.cellIterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    // Checking the cell type and format
                    // accordingly
                    switch (cell.getCellType()) {

                        // Case 1
                        case Cell.CELL_TYPE_NUMERIC:

                            if (cell.getRowIndex() == 0) {

                            } else {
                                System.out.print(cell.getNumericCellValue() + "  ");

                                if ((cell.getColumnIndex() == 3) && (ColumnsNames.get(3).equals("Devices"))) {

                                    ReadingList.get(cCount).setDevices((int) cell.getNumericCellValue());
                                }
                            }

                            break;

                        // Case 2
                        case Cell.CELL_TYPE_STRING:
                            if (cell.getRowIndex() == 0) {
                                ColumnsNames.add(cell.getStringCellValue());

                            } else {
                                System.out.print(cell.getStringCellValue() + "  ");

                                if ((cell.getColumnIndex() == 0) && (ColumnsNames.get(0).equals("Code"))) {

                                    ReadingList.add(new Key());
                                    cCount++;
                                    ReadingList.get(cCount).setCode(cell.getStringCellValue());
                                }

                                if ((cell.getColumnIndex() == 1) && (ColumnsNames.get(1).equals("Starting Date"))) {
                                    System.out.println("BLINKAS");
                                    ReadingList.get(cCount).setStartingDate(new SimpleDateFormat("yy/MM/dd").parse(cell.getStringCellValue()));
                                }

                                if ((cell.getColumnIndex() == 2) && (ColumnsNames.get(2).equals("Expiration Date"))) {
                                    System.out.println("BLINKAS2");

                                    if (!cell.getStringCellValue().equals("Unlimited")) {
                                        ReadingList.get(cCount).setExpDate(new SimpleDateFormat("yy/MM/dd").parse(cell.getStringCellValue()));
                                    }

                                }

                            }
                            break;
                    }
                }

                System.out.println("");
            }

            // Closing file output streams
            file.close();
        } // Catch block to handle exceptions
        catch (Exception e) {

            throw e;
            // Display the exception along with line number
            // using printStackTrace() method
            
            //e.printStackTrace();
        }

        System.out.println("************************************");
        for (int i = 0; i < ReadingList.size(); i++) {
            System.out.println(ReadingList.get(i).getCode() + "   " + ReadingList.get(i).getStartingDate() + "   " + ReadingList.get(i).getExpDate() + "   " + ReadingList.get(i).getDevices());

            Date d1 = new java.util.Date(); // Now Date 
            Date d2;

            if (ReadingList.get(i).getExpDate() != null) {
                d2 = ReadingList.get(i).getExpDate();

                long diff = d2.getTime() - d1.getTime();

                System.out.println("Difference between  " + d1 + " and " + d2 + " is "
                        + (diff / (1000 * 60 * 60 * 24)) + " days.");

                if ((int) (diff / (1000 * 60 * 60 * 24)) < 0) {
                    System.out.println("Expired");
                    ReadingList.get(i).setStatus("Expired");
                }

            } else {
                System.out.println("Unlimted");
                ReadingList.get(i).setStatus("Unlimted");
            }

        }

        return ReadingList;

    }

    public void ExcelWrite(String path) throws FileNotFoundException, IOException {
        // workbook object
        System.out.println("Path: " + path);

        XSSFWorkbook workbook = new XSSFWorkbook();

        // spreadsheet object
        XSSFSheet spreadsheet
                = workbook.createSheet(" Student Data ");

        // creating a row object
        XSSFRow row;

        // This data needs to be written (Object[])
        Map<String, Object[]> studentData
                = new TreeMap<String, Object[]>();

        studentData.put(
                "1",
                new Object[]{"Code", "Starting Date", "Expiration Date", "Devices"});

        /*
        studentData.put("2", new Object[] { "128", "Aditya",
                                            "2nd year" });
  
        studentData.put(
            "3",
            new Object[] { "129", "Narayana", "2nd year" });
  
        studentData.put("4", new Object[] { "130", "Mohan",
                                            "2nd year" });
  
        studentData.put("5", new Object[] { "131", "Radha",
                                            "2nd year" });
  
        
        studentData.put("6", new Object[] { "132", "Gopal",
                                            "2nd year" });
         */
        Set<String> keyid = studentData.keySet();

        int rowid = 0;

        // writing the data into the sheets...
        for (String key : keyid) {

            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = studentData.get(key);
            int cellid = 0;

            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);

            }
        }

        spreadsheet.setColumnWidth(0, 7500);
        spreadsheet.setColumnWidth(1, 7500);
        spreadsheet.setColumnWidth(2, 7500);

        // .xlsx is the format for Excel Sheets...
        // writing the workbook into the file...
        File f = new File(path);

        if (!f.exists()) {
            FileOutputStream out = new FileOutputStream(f);
            workbook.write(out);
            out.close();
            JOptionPane.showMessageDialog(null, "Created successfully");
        } else {
            JOptionPane.showMessageDialog(null, "File is already exists.");
        }

    }

    public void ExcelUpdate(String path, ArrayList<Key> keys, boolean Unlimited) throws FileNotFoundException, IOException, InvalidFormatException { // Updater
        // Creating file object of existing excel file
        File xlsxFile = new File(path);

        //New students records to update in excel file
        /*
        Object[][] newStudents = {
                {"Rakesh sharma", "New Delhi", "rakesh.sharma@gmail.com", 22},
                {"Thomas Hardy", "London", "thomas.hardy@gmail.com", 25}
        };
         */
        Object[][] newStudents = new Object[keys.size()][4];

        System.out.println("FARCRY");
        if (!Unlimited) {
            for (int i = 0; i < keys.size(); i++) {

                SimpleDateFormat DateFormat1 = new SimpleDateFormat("yy/MM/dd");
                SimpleDateFormat DateFormat2 = new SimpleDateFormat("yy/MM/dd");

                String curr_date1 = DateFormat1.format(keys.get(i).getStartingDate().getTime());
                String curr_date2 = DateFormat2.format(keys.get(i).getExpDate().getTime());

                newStudents[i][0] = keys.get(i).getCode();
                newStudents[i][1] = curr_date1;
                newStudents[i][2] = curr_date2;
                newStudents[i][3] = keys.get(i).getDevices();
            }
        } else {
            Date d1 = new java.util.Date(); // Now Date 
            SimpleDateFormat DateFormat = new SimpleDateFormat("yy/MM/dd");
            String curr_date = DateFormat.format(d1.getTime());

            for (int i = 0; i < keys.size(); i++) {

                newStudents[i][0] = keys.get(i).getCode();
                newStudents[i][1] = curr_date;
                newStudents[i][2] = "Unlimited";
                newStudents[i][3] = keys.get(i).getDevices();
            }

        }

        System.out.println("FARCRY2");
        try {
            //Creating input stream
            FileInputStream inputStream = new FileInputStream(xlsxFile);

            //Creating workbook from input stream
            Workbook workbook = WorkbookFactory.create(inputStream);

            //Reading first sheet of excel file
            Sheet sheet = workbook.getSheetAt(0);

            //Getting the count of existing records
            int rowCount = sheet.getLastRowNum();

            //Iterating new students to update
            for (Object[] student : newStudents) {

                //Creating new row from the next row count
                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                //Iterating student informations
                for (Object info : student) {

                    //Creating new cell and setting the value
                    Cell cell = row.createCell(columnCount++);
                    if (info instanceof String) {
                        cell.setCellValue((String) info);
                    } else if (info instanceof Integer) {
                        cell.setCellValue((Integer) info);
                    }
                }
            }
            //Close input stream
            inputStream.close();

            //Crating output stream and writing the updated workbook
            FileOutputStream os = new FileOutputStream(xlsxFile);
            workbook.write(os);

            //Close the workbook and output stream
            workbook.close();
            os.close();

            System.out.println("Excel file has been updated successfully.");

        } catch (EncryptedDocumentException | IOException e) {
            System.err.println("Exception while updating an existing excel file.");
            e.printStackTrace();
        }
    }

}
