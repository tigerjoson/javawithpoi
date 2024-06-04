package myexcel;
//ref https://stackoverflow.com/questions/1072561/how-can-i-read-numeric-strings-in-excel-cells-as-string-not-numbers
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class exceltocsv {

	public static void main(String[] args) throws Exception {
		DataFormatter poiFormatter = new DataFormatter();
		
		
		InputStream inputStream = new FileInputStream("....xlsx");
		InputStreamReader isr = new InputStreamReader(inputStream);
		
		OutputStream outputStream = new FileOutputStream("....csv");
		OutputStreamWriter osr = new OutputStreamWriter(outputStream);
		BufferedWriter bWriter = new BufferedWriter(osr);
		
		new BufferedReader(isr);
		XSSFWorkbook oriworkbook = new XSSFWorkbook(inputStream);
		 DateFormat.getDateInstance(0, Locale.TAIWAN);
		XSSFSheet ori_recordsheet = oriworkbook.getSheetAt(0);
	
		//set boundary
		for(int i=5;i<=1204;i++) {
			//System.out.print(oriworkbook.getSheetAt(0).getRow(i).getCell(0).toString());
			String departnameString = oriworkbook.getSheetAt(0).getRow(i).getCell(1).toString();
			String Brand_name =  oriworkbook.getSheetAt(0).getRow(i).getCell(2).toString();
			
			XSSFCell casenumbercCell =  oriworkbook.getSheetAt(0).getRow(i).getCell(3);
			String casenumberString = poiFormatter.formatCellValue(casenumbercCell);
			
			String hardiskserialnumber =  poiFormatter.formatCellValue(oriworkbook.getSheetAt(0).getRow(i).getCell(4));
			
			
			Date date =  oriworkbook.getSheetAt(0).getRow(i).getCell(9).getDateCellValue();
			//String dateString = poiFormatter.formatCellValue( oriworkbook.getSheetAt(0).getRow(i).getCell(9));
			//System.out.println(dateString);
			String recordString = departnameString+","+Brand_name+","+casenumberString+","+hardiskserialnumber+","+date.toString();
			//System.out.println(recordString);
			bWriter.write(recordString);
			bWriter.newLine();
		}
		System.out.println("fin~!");
		bWriter.close();
		osr.close();
		outputStream.close();
		isr.close();
		inputStream.close();
		
		
		
		
		/*
		for (int i=start_index_of_row;i<end_index_of_row;i++) {
			for(int ci=1;ci<=8;ci++) {
				String string= oriworkbook.getSheetAt(0).getRow(i).getCell(ci).toString();
				System.out.println(string);
			}
			
			
		} */
		
	

	}
}
