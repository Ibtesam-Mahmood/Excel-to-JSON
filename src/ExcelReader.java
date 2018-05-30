import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

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
		parseExcel();
		writer.WriteJSON();
	}
	
	private void parseExcel() {

		for(int n = 0; n < files.length; n++) {


		    if(isExcel(files[n]))
		        parseSheet(files[n]);
//            try{
//
//                Workbook w =  Workbook.getWorkbook(files[n]);
//                sheet = w.getSheet(0);
//
//            } catch(BiffException | IOException e) { e.printStackTrace(); return null; }
//
//            for (int i = 0; i < sheet.getRows(); i++) {
//                int size = checkRowSize(i, sheet);
//
//
//                parse[n] += sheet.getCell(0, i).getContents() + ":";
//
//                for(int j = 0; j < size; j++) {
//                    parse[n] += sheet.getCell(j+1, i).getContents();
//                    if(j != size-1)
//                        parse[n] += ",";
//                }
//
//
//                if(i != sheet.getRows() - 1)
//                    parse[n] += ";";
//            }

		}

	}

	//Parses the Main sheet in the given file
	private void parseSheet(File file){

        Sheet sheet;
        JSONObject object;

        try{

                Workbook w =  Workbook.getWorkbook(file);
                sheet = w.getSheet(0);

        } catch(BiffException | IOException e) { e.printStackTrace(); return; }

        object =  createJSONObject(0, sheet.getRows(), sheet);

    }

    //Creates a JSON object with the parameters defined between the start and end
    private JSONObject createJSONObject(int startRow, int endRow, Sheet sheet){

        JSONObject object =  new JSONObject();

        for (int i = 0; i < sheet.getRows(); i++) {
            int size = checkRowSize(i, sheet);
            String name = sheet.getCell(0, i).getContents();

            if(size == 2){
                if(name == "}") //Ending line of an object
                    continue;
                else if(name == "{"){ //Starting of an object

                }
                else{ //Simple value
                    String value = sheet.getCell(1, i).getContents();
                    object.put(name, value);
                }


            }
            else{ //Create an array
                JSONArray array = createJSONArray(i, sheet);
                object.put(name, array);
            }


        }

        return object;

    }

    //Creates a JSON array out of a given row
    private JSONArray createJSONArray(int row, Sheet sheet){

	    JSONArray array =  new JSONArray();

	    for (int i = 1; ; i++){

	        String contents = sheet.getCell(i, row).getContents();

	        if(contents == "")
	            break;

	        array.add(contents);

        }

        return array;

    }

	//Determines if a file is an xls file
    //xlsx files are not compatible
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

	//Determines the size of the row on a given sheet
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
