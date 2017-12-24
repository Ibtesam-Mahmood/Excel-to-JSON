
public class Main {

	public static void main(String[] args) {
		ExcelReader reader =  new ExcelReader();
		System.out.println( reader.parseExcel()[0] );
	}
	
}
