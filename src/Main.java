import java.nio.file.Path;
import java.nio.file.Paths;

public class Main {

	Path initialPath = Paths.get("files\\Base").toAbsolutePath();
	Path endingPath = Paths.get("files\\Final").toAbsolutePath();

	public static void main(String[] args) {
		ExcelReader reader =  new ExcelReader();
		reader.convert();
	}
	
}
