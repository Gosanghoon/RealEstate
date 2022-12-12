package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class JMMultipleratioWrite {

    public void writeExcel(HashMap<Integer, ArrayList<String>> hashMap, String filePath, String outFileNm){
        //Empty Workbook create
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Empty Sheet create
        XSSFSheet sheet = workbook.createSheet("RealEstate Data");

        // DATA에서 KETSET을 가져온다. 이 SET 값들을 조회하면서 데이터들을 SHEET에 입력한다.
        Set<Integer> keyset = hashMap.keySet();
        int rownum = 0;

        // 알아야할 점, TreeMap을 통해 생성된 ketSet은 for 조회시, 키값이 오름차순으로 조회된다.
        for(Integer key : keyset){
            Row row = sheet.createRow(rownum++);
            ArrayList<String> objArr = hashMap.get(key);
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

        try (FileOutputStream out = new FileOutputStream(new File(filePath, outFileNm))){
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
