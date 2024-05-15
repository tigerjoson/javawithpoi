package word;

import java.io.BufferedReader;
import java.io.BufferedWriter;
//ref chatgpt
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ReadWordTablefromhncbform {
	public static void main(String[] args) {

		try {
			// Specify the path to your DOCX file
			String filePath = "docx file path";
			InputStream inputStream = new FileInputStream(filePath);
			InputStreamReader isr = new InputStreamReader(inputStream);
			BufferedReader bReader = new BufferedReader(isr);
			FileWriter fWriter = new FileWriter("write to csv");
			BufferedWriter bWriter = new BufferedWriter(fWriter);
			// Create a FileInputStream to read the DOCX file
			FileInputStream fis = new FileInputStream(filePath);

			// Create an XWPFDocument from the FileInputStream
			XWPFDocument document = new XWPFDocument(fis);
			List<XWPFTable> tablelist = document.getTables();
			XWPFTable table = tablelist.get(0);
			//set table boundary
			for(int i=2;i<11;i++) {
				for (int j=0;j<4;j++) {
					bWriter.write( table.getRow(i).getCell(j).getText()+",");
				}
				bWriter.newLine();
			}
			
			
			System.out.println("fin~!");
			bReader.close();
			inputStream.close();
			isr.close();
			bWriter.close();
			fWriter.close();
			// Close the FileInputStream
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
