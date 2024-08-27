package mypptx;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
//ref : win11 copilot
public class printtablefrompptx {

	public static void main(String[] args) throws Exception {
		FileInputStream fis = new FileInputStream("C:\\Users\\tiger\\OneDrive\\桌面\\參考資料\\hncb綜合系統科\\code_riview_P073002_SnY家事事件查詢\\uipath_P073002_SnY家事事件查詢.pptx");
		XMLSlideShow pptx = new XMLSlideShow(fis);

		for (XSLFSlide slide : pptx.getSlides()) {
		    for (XSLFShape shape : slide) {
		        if (shape instanceof XSLFTable) {
		            XSLFTable table = (XSLFTable) shape;
		            // Process table content here (e.g., read cell values)
		            // Example: String cellValue = table.getCell(0, 0).getText();
		            String cellValue = table.getCell(0, 0).getText();
		            System.out.println(cellValue);
		        }
		    }
		}

	}

}
