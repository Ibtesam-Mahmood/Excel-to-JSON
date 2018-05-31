import java.nio.file.Path;
import java.nio.file.Paths;

public class Main {


	public static void main(String[] args) {

		final Path initialPath = Paths.get("files\\Base").toAbsolutePath(); //Starting path for excel files
		final Path endingPath = Paths.get("files\\Final").toAbsolutePath(); //Ending path for JSON files

		ExcelToJSONParser reader =  new ExcelToJSONParser(initialPath, endingPath);
	}
	
}
