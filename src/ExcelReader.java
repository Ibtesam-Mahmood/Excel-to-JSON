import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelReader{

	
	private final Path initialPath = Paths.get("files\\Base").toAbsolutePath();
	private final Path finalPath =  Paths.get("files\\Final").toAbsolutePath();
	
	private File files;
	
	public ExcelReader() {
		files = initialPath.toFile().listFiles()[0];
	}
	
	public String parseExcel() {
		
		String parse = "";
		
		Sheet sheet;
		
		try{
			
			Workbook w =  Workbook.getWorkbook(files);
			sheet = w.getSheet(0);
			
		} catch(BiffException | IOException e) { e.printStackTrace(); return null; }
		
		for (int i = 0; i < sheet.getRows(); i++) {
			parse += sheet.getCell(0, i).getContents() + ":" + sheet.getCell(1, i).getContents() + ";\n";
		}
		
		
		return parse;
		
	}
	
	public boolean isExcel(File file) {
		String name = file.getName();
		String format = name.substring(name.indexOf("."), name.length());
		if(format.equalsIgnoreCase("xls")) {
			return true;
		}
		else {
			System.out.println(name + " is not the correct format for this program");
			return false;
		}
		
		
	}
	
	
	
}
