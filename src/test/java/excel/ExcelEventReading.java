package excel;

import vip.ipav.poi.excel.reader.Excel2003Reader;
import vip.ipav.poi.excel.reader.Excel2007Reader;
import org.junit.Test;

import java.util.Iterator;

/**
 * Created by 89003522 on 2017/11/2.
 */
public class ExcelEventReading {

    @Test
    public void say(){
        System.out.println("hello world!");
    }

    @Test
    public void readExcel2007() throws Exception{
        //读取文件
        Excel2007Reader excel2007Reader = new Excel2007Reader(1,"D:/test.xlsx");

        excel2007Reader.processOneSheet(1);
        //excel2007Reader.processOneSheet(2);

        //excel2007Reader.processAllSheets();

        Iterator it = excel2007Reader.getAllValueList().iterator();
        while (it.hasNext()){
            System.out.println(it.next());
        }
    }

    @Test
    public void readExcel2003() throws Exception{
       Excel2003Reader excel2003Reader = new Excel2003Reader(1,"d:/test.xls");
//       excel2003Reader.processAllSheets();
        excel2003Reader.processOneSheet(1);
       Iterator it = excel2003Reader.getAllValueList().iterator();
       while (it.hasNext()) {
            System.out.println(it.next());
       }
    }
}

