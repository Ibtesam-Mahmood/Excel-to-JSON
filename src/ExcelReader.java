import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelReader{

	
	private Path initialPath;

	private JSONWriter writer;
	
	private File[] files;
	
	public ExcelReader(Path init, Path fin)
	{
	    this.initialPath = init;
		this.files = initialPath.toFile().listFiles();
		this.writer = new JSONWriter(fin, this.files);
	}
	
	public void convert() {
		String[] JSONObjects = parseExcel();
		writer.WriteJSON();
	}
	
	private String[] parseExcel() {
		
		String[] parse = new String[files.length];
		
		for (int i = 0; i < parse.length; i++) {
			parse[i] = "";
		}
		
		Sheet sheet;
		
		for(int n = 0; n < files.length; n++) {
			
			try{
				
				Workbook w =  Workbook.getWorkbook(files[n]);
				sheet = w.getSheet(0);
				
			} catch(BiffException | IOException e) { e.printStackTrace(); return null; }
			
			for (int i = 0; i < sheet.getRows(); i++) {
				int size = checkRowSize(i, sheet);
				
				
				parse[n] += sheet.getCell(0, i).getContents() + ":";
				
				for(int j = 0; j < size; j++) {
					parse[n] += sheet.getCell(j+1, i).getContents();
					if(j != size-1)
						parse[n] += ",";
				}
					
				
				if(i != sheet.getRows() - 1)
					parse[n] += ";";
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
