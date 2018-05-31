import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

public class ExcelToJSONParser {

	
	private Path initialPath;
	private Path finalPath;
	
	private File[] files;
	
	public ExcelToJSONParser(Path init, Path fin)
	{
	    this.initialPath = init;
	    this.finalPath = fin;
		this.files = initialPath.toFile().listFiles();
	}

	//Loops through the files within the initial directory and parses them
	public void parse() {

		for(int n = 0; n < files.length; n++) {

		    //Checks if the file is and xls file
		    if(isExcel(files[n]))
		        parseSheet(files[n]);

		}

	}

	private void writeJSON(JSONObject object, File file){

        String dir = finalPath.toString() + "\\" +  JSONName(file);

        try( FileWriter writer = new FileWriter(dir) ) {

            writer.write(object.toString());

        } catch (IOException e) { e.printStackTrace(); return;}

    }

	//Parses the Main sheet in the given file
	private void parseSheet(File file){

        Sheet sheet;
        JSONObject object;

        try{
                //Obtains the main sheet
                Workbook w =  Workbook.getWorkbook(file);
                sheet = w.getSheet(0);

        } catch(BiffException | IOException e) { e.printStackTrace(); return; }

        //Creates the JSON object and writes it
        object =  createJSONObject(0, sheet.getRows(), sheet);
        writeJSON(object, file);

    }

    //Creates a JSON object with the parameters defined between the start and end
    private JSONObject createJSONObject(int startRow, int endRow, Sheet sheet){

        JSONObject object =  new JSONObject();

        for (int i = 0; i < sheet.getRows(); i++) {
            int size = checkRowSize(i, sheet);
            String name = sheet.getCell(0, i).getContents();

            if(size == 2){
                if(name == "}") //Ending line of an object
                    continue; //Skip to the next line
                else if(name == "{"){ //Starting of an object

                }
                else{ //Simple value added to the parent json object
                    String value = sheet.getCell(1, i).getContents();
                    object.put(name, value);
                }


            }
            else{ //Create an array and add it to the parent json object
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

    //Converts the file name into a JSON file name
    private String JSONName(File file) {
        String orName = file.getName();
        return orName.substring(0, orName.indexOf(".") ) + ".Json";
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
