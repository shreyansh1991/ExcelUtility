package tests;

import java.io.File;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import excel.ReadDataFromExcelSheet;

public class TestDemo {
	@Test(dataProvider = "Passing List Of Maps")
	public void test(Map<String,String> map) {
		System.out.println("Hello..."+map.get("shreyanshData3"));
	}

	@DataProvider(name = "Passing List Of Maps")
	public Iterator<Object[]> createDataforTest3() {
		String excellocation = System.getProperty("user.dir")+File.separator+"shreyansh.xlsx";
		String sheetName = "Sheet1";
		
		ReadDataFromExcelSheet excelSheet = new ReadDataFromExcelSheet(excellocation);
		List<Map<String, String>> lom = excelSheet.getExcelData(excellocation, sheetName);
		
		Collection<Object[]> dp = new ArrayList<Object[]>();
		for (Map<String, String> map : lom) {
			dp.add(new Object[] { map });
		}
		return dp.iterator();
	}

}
