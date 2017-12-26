import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.json.simple.JSONArray;
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

				JSONObject jObj = extractFromText(parsedExcel[i]);
				writer.write(jObj.toString());
				
			} catch (IOException e) { e.printStackTrace(); continue;}
			
		
			
		}
		
		
	}
	
	
	private JSONObject extractFromText(String text) {
		
		JSONObject jObj = new JSONObject();
		
		String[] components = text.split(";");
		
		for (int i = 0; i < components.length; i++) {
			
			String[] variables = components[i].split(":");
			
			String name = variables[0];
			String[] values = variables[1].split(",");
			
			if(values.length == 1) {
				jObj.put(name, values[0]);
			}
			else {
				
				JSONArray array =  new JSONArray();
				
				for (int j = 0; j < values.length; j++) {
					
					array.add(values[j]);
					
				}
				
				jObj.put(name, array);
				
			}
			
		}
		
		return jObj;
	}
	
	
	private String JSONName(File file) {
		String orName = file.getName();
		return orName.substring(0, orName.indexOf(".") ) + ".Json";
	}
	
}
