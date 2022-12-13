package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class JMMultipleratioWrite {

    public void writeExcel(HashMap<Integer, ArrayList<String>> hashMap, HashMap<Integer, ArrayList<String>> maemaeMap, String filePath, String outFileNm){
        //Empty Workbook create
        XSSFWorkbook workbook = new XSSFWorkbook();

        for(int i = 0 ; i < 2; i++){

            //Empty Sheet create
            String sheetName = i == 0 ? "전세" : "매매";
            XSSFSheet sheet = workbook.createSheet(sheetName);

            // DATA에서 KETSET을 가져온다. 이 SET 값들을 조회하면서 데이터들을 SHEET에 입력한다.
            Set<Integer> keyset = i == 0 ? hashMap.keySet() : maemaeMap.keySet();
            //Set<Integer> keyset = hashMap.keySet();
            int rownum = 0;

            for(Integer key : keyset){
                Row row = sheet.createRow(rownum++);
                ArrayList<String> objArr = i == 0 ? hashMap.get(key) : maemaeMap.get(key);
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


}
