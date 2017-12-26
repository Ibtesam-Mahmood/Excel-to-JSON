import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.json.simple.JSONObject;

public class JSONWriter {

	private final Path finalPath =  Paths.get("files\\Final").toAbsolutePath();
	
	private String[] parsedExcel;
	private File[] files;
	
	public JSONWriter(String[] parsedExcel, File[] files) {
		this.parsedExcel = parsedExcel;
		this.files = files;
	}
	
	public void WriteJSON() {
		
		for (int i = 0; i < files.length; i++) {
			
			String dir = finalPath.toString() + "\\" +  JSONName(files[i]);
			
			try( FileWriter writer = new FileWriter(dir) ) {

				
				
			} catch (IOException e) { e.printStackTrace(); continue;}
			
		
			
		}
		
		
	}
	
	
	private JSONObject extractFromText(String text) {
		
		JSONObject jObj = new JSONObject();
		
		
		
		return jObj;
	}
	
	
	private String JSONName(File file) {
		String orName = file.getName();
		return orName.substring(0, orName.indexOf(".") ) + ".txt";
	}
	
}
