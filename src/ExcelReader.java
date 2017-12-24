import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import jxl.Cell;
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
			int size = checkRowSize(i, sheet);
			
			if(size == 1)
				parse += sheet.getCell(0, i).getContents() + ":" + sheet.getCell(1, i).getContents() + ";\n";
			else if(size > 1) {
				parse += sheet.getCell(0, i).getContents() + ":[";
				for(int j = 0; j < size; j++) {
					parse += sheet.getCell(j+1, i).getContents();
					if(j != size-1)
						parse += ",";
				}
				parse += "];\n";
			}
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
	
	private int checkRowSize(int row, Sheet sheet) {
		int n = 0;
		
		for (int i = 1; i < sheet.getColumns(); i++) {
			
			Cell temp =  sheet.getCell(i, row);
			
			if(temp.getContents() == "")
				break;
			else
				n++;
			
		}
		
		return n;
	}
	
	
	
}
