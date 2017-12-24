import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.file.Path;
import java.nio.file.Paths;

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
			
			
			try {
				
				FileOutputStream f = new FileOutputStream(finalPath.toString() + "\\" +  JSONName(files[i]) );
				Writer writer = new BufferedWriter( new OutputStreamWriter(f) );
				writer.write("{");
				
				String[] indexObjects = parsedExcel[i].split(";");
				for (int j = 0; j < indexObjects.length; j++) {
					writer.write(indexObjects[j]);
					if(j != indexObjects.length - 1)
						writer.write(",");
				}
				
				writer.write("}");
				
			} catch (IOException e) { e.printStackTrace(); continue; }
			
		
			
		}
		
		
	}
	
	
	private String JSONName(File file) {
		String orName = file.getName();
		return orName.substring(0, orName.indexOf(".") ) + ".txt";
	}
	
}
