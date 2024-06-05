package myexcel;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JToolBar;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import my_tool.Getfileproperties;

public class print_xlsx_info implements ActionListener {
	public JScrollPane jScrollPane;
	public JPanel corner;
	public JToolBar jToolBar, blank_jJToolBar;
	public JFileChooser jFileChooser;
	public FileNameExtensionFilter filter;
	public JButton openButton;
	public JFrame basic_frame;
	public Getfileproperties getfileproperties;

	public print_xlsx_info() {
		basic_frame = new JFrame("check excel");
		basic_frame.setBounds(100, 100, 450, 300);
		basic_frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		jToolBar = new JToolBar();
		basic_frame.add(jToolBar, BorderLayout.NORTH);

		openButton = new JButton("open xlsx file");
		jToolBar.add(openButton);
		openButton.addActionListener(this);

		basic_frame.setVisible(true);

	}

	public static void main(String[] args) {

		print_xlsx_info print_xlsx_info = new print_xlsx_info();

	}

	@Override
	public void actionPerformed(ActionEvent e) {

		jFileChooser = new JFileChooser("C:\\...");
		//jFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
		filter = new FileNameExtensionFilter("office 2013 above", "xlsx");
		//filter = new FileNameExtensionFilter();
		jFileChooser.setFileFilter(filter);
		
		int return_value = jFileChooser.showOpenDialog(basic_frame);
		
		if (return_value == JFileChooser.APPROVE_OPTION) {
			// System.out.println("return_value "+ return_value);

			File selectedfile = jFileChooser.getSelectedFile();
			String filepath = selectedfile.getAbsolutePath();
			System.out.println(filepath);
			try {
				InputStream inputStream_mylist = new FileInputStream(selectedfile);
				XSSFWorkbook workbook = new XSSFWorkbook(inputStream_mylist);
				System.out.println("ver="+workbook.getSpreadsheetVersion());
				System.out.println("1st sheet name ="+workbook.getSheetName(0));
				System.out.println("Number Of Sheets ="+workbook.getNumberOfSheets());

			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			

			jScrollPane = new JScrollPane();
			// jScrollPane.setColumnHeaderView(jToolBar);
			jScrollPane.setCorner(JScrollPane.UPPER_LEFT_CORNER, corner);
			jScrollPane.setCorner(JScrollPane.UPPER_RIGHT_CORNER, corner);
			jScrollPane.setCorner(JScrollPane.LOWER_LEFT_CORNER, corner);
			jScrollPane.setCorner(JScrollPane.LOWER_RIGHT_CORNER, corner);

		}

	}
}
