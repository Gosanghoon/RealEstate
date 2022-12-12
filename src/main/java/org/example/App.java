package org.example;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;


/**
 * Hello world!
 *
 */
public class App 
{
    public static String filepath = "C:\\poi_temp";
    public static String fileNm   = "poi_making_file_test.xlsx";

    public static void main( String[] args )
    {
        //Empty Workbook create
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Empty Sheet create
        XSSFSheet sheet = workbook.createSheet("RealEstate Data");

        Map<String, Object[]> data = new TreeMap<>();
        data.put("1", new Object[]{"ID","NAME","NUMBER"});
        data.put("2", new Object[]{"1","TEST","NUMBER"});
        data.put("3", new Object[]{"2","TEST","NUMBER"});
        data.put("4", new Object[]{"3","TEST","NUMBER"});
        data.put("5", new Object[]{"4","tEST","NUMBER"});

        // DATA에서 KETSET을 가져온다. 이 SET 값들을 조회하면서 데이터들을 SHEET에 입력한다.
        Set<String> keyset = data.keySet();
        int rownum = 0;

        // 알아야할 점, TreeMap을 통해 생성된 ketSet은 for 조회시, 키값이 오름차순으로 조회된다.
        for(String key : keyset){
            Row row = sheet.createRow(rownum++);
            Object[] objArr = data.get(key);
            int cellnum = 0;
            for(Object obj : objArr){
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String){
                    cell.setCellValue((String)obj);
                } else if (obj instanceof Integer){
                    cell.setCellValue((Integer)obj);
                }
            }
        }

        try (FileOutputStream out = new FileOutputStream(new File(filepath, fileNm))){
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
