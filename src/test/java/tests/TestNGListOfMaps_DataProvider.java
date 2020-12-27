package tests;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
import java.util.*;

public class TestNGListOfMaps_DataProvider {
    @DataProvider(name = "Passing List Of Maps")
    public Iterator<Object[]> createDataforTest3() {
        List<Map<String,String>> lom = new ArrayList<Map<String,String>>();
        Map<String,String> map1 = new HashMap<String,String>();
        Map<String,String> map2 = new HashMap<String,String>();
        map1.put("name","Amit");
        map2.put("name","Sarbjit");
        lom.add(map1);
        lom.add(map2);
        Collection<Object[]> dp = new ArrayList<Object[]>();
        for(Map<String,String> map:lom){
            dp.add(new Object[]{map});
        }
        return dp.iterator();
    }

    @Test(dataProvider = "Passing List Of Maps")
    public void test1(Map<String,String> map){
        System.out.println("Value in first Map:" + map.get("name"));
    }
}
