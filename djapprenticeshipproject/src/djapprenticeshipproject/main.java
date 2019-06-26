package djapprenticeshipproject;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class main {

	public static void main(String[] args) throws InvalidFormatException {
		// all the main methods will go here later
		//msExcel.appendNewRow();
		String[] rowToAppend = new String[]{"VM152", "AnotherBot", "ABot Live","100089","fakepassword89","", "OE Robotics", "AVM1089","LiveCHS","RTRA",""};
		msExcel.appendInputRow(rowToAppend);
		String[] nextRowToAppend = new String[]{"VM154", "AnotherBot", "ABot Live","100094","fakepassword94","", "OE Robotics", "AVM1094","LiveCHS","RTRA",""};
		msExcel.appendInputRow(nextRowToAppend);
		String[] anotherRowToAppend = new String[]{"VM153", "AnotherBot", "ABot Live","1000890","fakepassword90","", "OE Robotics", "AVM1090","LiveCHS","RTRA",""};
		msExcel.appendInputRow(anotherRowToAppend);;
		System.out.println("done");
	}

}
