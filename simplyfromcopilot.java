package word;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JToolBar;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class Simplydocxfrmai implements ActionListener {

	public JScrollPane jScrollPane;
	public JPanel corner;
	public JToolBar jToolBar, blank_jJToolBar;
	public JFileChooser jFileChooser;
	public FileNameExtensionFilter filter;
	public JButton openButton;
	public JFrame basic_frame;

	static final String wanttoremoveString_reg = "**";

	public Simplydocxfrmai() {
		basic_frame = new JFrame("simply");
		basic_frame.setBounds(100, 100, 450, 300);
		basic_frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		jToolBar = new JToolBar();
		basic_frame.add(jToolBar, BorderLayout.NORTH);

		openButton = new JButton("open docx file");
		jToolBar.add(openButton);
		openButton.addActionListener(this);

		basic_frame.setVisible(true);

	}

	public static void main(String[] args) {

		Simplydocxfrmai s = new Simplydocxfrmai();

	}

	@Override
	public void actionPerformed(ActionEvent e) {

		jFileChooser = new JFileChooser("your default path");
		// jFileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

		jFileChooser.setFileFilter(new FileNameExtensionFilter("docx files", "docx"));
		int return_value = jFileChooser.showOpenDialog(basic_frame);

		if (return_value == JFileChooser.APPROVE_OPTION) {
			// System.out.println("return_value " + return_value);
			try {
				File file = new File(jFileChooser.getSelectedFile().getPath());
				String simplydocx = file.getPath().replaceAll(".docx$", "simply.docx");
				FileOutputStream out = new FileOutputStream(simplydocx);
				FileInputStream fis = new FileInputStream(file.getAbsolutePath());
				XWPFDocument outdocument = new XWPFDocument();
				XWPFDocument inputdocument = new XWPFDocument(fis);

				List<XWPFParagraph> paragraphs = inputdocument.getParagraphs();
				XWPFParagraph paragraph = outdocument.createParagraph();
				for (XWPFParagraph para : paragraphs) {
					String lineString = para.getText();
					String newString= lineString.replaceAll("\\**", "").replaceAll("來源: 與 Copilot 的交談\\D", "").replaceAll("如果你對其他\\D{1,}歡迎詢問！", "").replaceAll("Are you encountering a specific issue with a\\D{1,}? Maybe I can help troubleshoot it further!", "").replaceAll("#{3}", "");
//					paragraph.createRun().setText(newString);
//					paragraph.createRun().addBreak();
					XWPFParagraph paragraph1 = outdocument.createParagraph();
					paragraph1.createRun().setText(newString);
				}

				outdocument.write(out);
				out.close();
				System.exit(0);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

		}

		jScrollPane = new JScrollPane();
		// jScrollPane.setColumnHeaderView(jToolBar);
		jScrollPane.setCorner(JScrollPane.UPPER_LEFT_CORNER, corner);
		jScrollPane.setCorner(JScrollPane.UPPER_RIGHT_CORNER, corner);
		jScrollPane.setCorner(JScrollPane.LOWER_LEFT_CORNER, corner);
		jScrollPane.setCorner(JScrollPane.LOWER_RIGHT_CORNER, corner);

	}

}
