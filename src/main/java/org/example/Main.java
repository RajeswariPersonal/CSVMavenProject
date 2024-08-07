package org.example;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.util.ArrayList;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
public class Main {
    public static void main(String[] args) throws IOException {
        //TIP Press <shortcut actionId="ShowIntentionActions"/> with your caret at the highlighted text
        // to see how IntelliJ IDEA suggests fixing it.
        System.out.printf("welcome!");

        ReadAndWriteToExcel();
    }

    public static void ReadAndWriteToExcel() throws IOException
    {
        ArrayList<ArrayList<String>> allRowAndColData = null;
        ArrayList<String> oneRowData = null;
        String fName = "C:\\Users\\User\\Documents\\test.csv";
        String[] currentLine;
        CSVReader reader = null;
        reader = new CSVReader(new FileReader(fName));

        allRowAndColData = new ArrayList<ArrayList<String>>();
        while (true) {
            try {
                if (!((currentLine = reader.readNext()) != null)) break;
            } catch (CsvValidationException e) {
                throw new RuntimeException(e);
            }
            oneRowData = new ArrayList<String>();
            for (int j = 0; j < currentLine.length; j++) {
                oneRowData.add(currentLine[j]);
            }
            allRowAndColData.add(oneRowData);
            System.out.println();

        }

        try {
            HSSFWorkbook workBook = new HSSFWorkbook();
            HSSFSheet sheet = workBook.createSheet("sheet1");
            for (int i = 0; i < allRowAndColData.size(); i++) {
                ArrayList<String> ardata = (ArrayList<String>) allRowAndColData.get(i);
                HSSFRow row = sheet.createRow(0 + i);
                for (int k = 0; k < ardata.size(); k++) {
                    System.out.print(ardata.get(k));
                    HSSFCell cell = row.createCell(k);
                    cell.setCellValue(ardata.get(k).toString());
                }
                System.out.println();
            }
            FileOutputStream fileOutputStream =  new FileOutputStream("C:\\Users\\User\\Documents\\sample.xls");
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("File Exported Successfully in Excel file");
        } catch (Exception ex) {
        }
    }
}