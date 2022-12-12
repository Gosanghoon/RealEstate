package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Array;
import java.util.*;

public class JMMultiplesratioRead {

    public static String filePath = "C:\\poi_temp";
    public static String fileNm = "poi_reading_file_test.xlsx";
    public static String outFileNm = "ksh_write_file_test.xlsx";
    public static int restPrice = 0;
    public static int hundredMilion = 0;

    public static void main(String[] args){

        try (FileInputStream file = new FileInputStream(new File(filePath, fileNm))){

            //엑셀 파일로 Workbook instance를 생성한다.
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //workbook의 첫번째 sheet를 가져온다.
            XSSFSheet sheet = workbook.getSheetAt(0);

            HashMap<Integer, ArrayList<String>> resultMap = new HashMap<Integer, ArrayList<String>>();


            /*
                만약 특정 이름의 시트를 찾는다면 workbook.getSheet("찾는 시트의 이름");
                만약 모든 시트를 순회하고 싶다면
                for(Integer sheetNum : workbook.getNumberOfSheets()){
                    XSSFSheet sheet = workbook.getSheetAt(i);
                }
                아니면 Iterator<Sheet> s = workbook.iterator()를 사용해서 조회해도 좋다.
             */

            for (Row row : sheet){
                //각각의 행에 존재하는 모든 cell을 순회한다.
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();

                    // cell의 타입을 확인하고 값을 가져온다.
                    switch (cell.getCellType()){

                        case NUMERIC:
                            //System.out.print("타니?");
                            //getNumericCellValue 메서드는 기본으로 double형 반환
                            System.out.print(cell.getColumnIndex() + "\t");
                            System.out.print(cell.getNumericCellValue() + "\t");

                            if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 0) {
                                resultMap.put(cell.getRowIndex() -1, new ArrayList<String>());
                                resultMap.get(cell.getRowIndex() -1).add(cell.getStringCellValue());
                            }
                            break;

                        case STRING:

                            String targetPrice = cell.getStringCellValue();
                            String transePrice = "";
                            //System.out.print(targetPrice + "\t");
                            //System.out.print(cell.getRow());

                            if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 0) {
                                resultMap.put(cell.getRowIndex() -1, new ArrayList<String>());
                                resultMap.get(cell.getRowIndex() -1).add(cell.getStringCellValue());

                            }

                            if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 1){
                                System.out.print(targetPrice + "\t");
                                //1. 공백제거, 쉼표 제거
                                if (targetPrice.contains(" "))
                                    transePrice = targetPrice.replace(" ","");

                                if (targetPrice.contains(","))
                                    transePrice = transePrice.replace(",","");

                                //2. 천만원단위 가공
                                if (!transePrice.substring(transePrice.indexOf("억")+1, transePrice.length()).equals(""))
                                    restPrice = Integer.parseInt(transePrice.substring(transePrice.indexOf("억")+1, transePrice.length())) * 10000;

                                //3. 억단위 가공
                                hundredMilion = Integer.parseInt(targetPrice.substring(0, targetPrice.indexOf("억"))) * 100000000;

                                //4. 합치기
                                String finalPrice = String.valueOf(hundredMilion + restPrice);
                                System.out.print(finalPrice + "\t");
                                resultMap.get(cell.getRowIndex() -1).add(finalPrice);
                            }
                            break;
                        }
                    //addArr.clear();
                }
            }

            for(Integer i : resultMap.keySet()){
                System.out.println("[key]:" + i + " [value]:" + resultMap.get(i));
            }

            JMMultipleratioWrite jmMultipleratioWrite = new JMMultipleratioWrite();
            jmMultipleratioWrite.writeExcel(resultMap, filePath, outFileNm);
            System.out.println();

        } catch (Exception e){
            e.printStackTrace();
        }

    }


}



