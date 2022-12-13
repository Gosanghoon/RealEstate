package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class JMMultiplesratioRead {

    public static String filePath = "C:\\poi_temp";

    public static void main(String[] args){

        Scanner sc = new Scanner(System.in);

        System.out.println("읽을 파일명을 입력하세요.");
        String inputFileNm = sc.next();
        System.out.println("출력할 파일명을 입력하세요.");
        String outputFileNm = sc.next();
        System.out.println(inputFileNm + "   " + outputFileNm);

        try (FileInputStream file = new FileInputStream(new File(filePath, inputFileNm))){


            //엑셀 파일로 Workbook instance를 생성한다.
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            HashMap<Integer, ArrayList<String>> resultMap = new HashMap<Integer, ArrayList<String>>();
            HashMap<Integer, ArrayList<String>> maemaeMap = new HashMap<Integer, ArrayList<String>>();

            for(int i = 0; i < workbook.getNumberOfSheets(); i++) {

                //get sheet
                XSSFSheet sheet = workbook.getSheetAt(i);

            /*
                만약 특정 이름의 시트를 찾는다면 workbook.getSheet("찾는 시트의 이름");
                만약 모든 시트를 순회하고 싶다면
                for(Integer sheetNum : workbook.getNumberOfSheets()){
                    XSSFSheet sheet = workbook.getSheetAt(i);
                }
                아니면 Iterator<Sheet> s = workbook.iterator()를 사용해서 조회해도 좋다.
             */

                for (Row row : sheet) {
                    //각각의 행에 존재하는 모든 cell을 순회한다.
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();

                        // cell의 타입을 확인하고 값을 가져온다.
                        switch (cell.getCellType()) {

                            case NUMERIC:
                                System.out.print(cell.getColumnIndex() + "\t");
                                System.out.print(cell.getNumericCellValue() + "\t");

                                if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 0) {
                                    resultMap.put(cell.getRowIndex() - 1, new ArrayList<String>());
                                    resultMap.get(cell.getRowIndex() - 1).add(cell.getStringCellValue());
                                }
                                break;

                            case STRING:

                                String targetPrice = cell.getStringCellValue();
                                String transePrice = "";
                                int restPrice = 0;
                                int hundredMilion = 0;

                                if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 0) {
                                    if(sheet.getSheetName().equals("전세")){
                                        resultMap.put(cell.getRowIndex() - 1, new ArrayList<String>());
                                        resultMap.get(cell.getRowIndex() - 1).add(cell.getStringCellValue());
                                    } else if(sheet.getSheetName().equals("매매")){
                                        maemaeMap.put(cell.getRowIndex() - 1, new ArrayList<String>());
                                        maemaeMap.get(cell.getRowIndex() - 1).add(cell.getStringCellValue());
                                    }
                                }

                                if (cell.getRowIndex() > 0 && cell.getColumnIndex() == 1) {

                                    //1. 공백제거, 쉼표 제거
                                    if (targetPrice.contains(" "))
                                        transePrice = targetPrice.replace(" ", "");

                                    if (targetPrice.contains(","))
                                        transePrice = transePrice.replace(",", "");

                                    //2. 천만원단위 가공
                                    if (!transePrice.substring(transePrice.indexOf("억") + 1, transePrice.length()).equals(""))
                                        restPrice = Integer.parseInt(transePrice.substring(transePrice.indexOf("억") + 1, transePrice.length())) * 10000;

                                    //3. 억단위 가공
                                    hundredMilion = Integer.parseInt(targetPrice.substring(0, targetPrice.indexOf("억"))) * 100000000;

                                    //4. 합치기
                                    String finalPrice = String.valueOf(hundredMilion + restPrice);
                                    //System.out.print(finalPrice + "\t");
                                    if(sheet.getSheetName().equals("전세")){
                                        resultMap.get(cell.getRowIndex() - 1).add(finalPrice);
                                    } else if(sheet.getSheetName().equals("매매")){
                                        maemaeMap.get(cell.getRowIndex() - 1).add(finalPrice);
                                    }
                                }
                                break;
                        }
                    }
                }
            }

            for(Integer i : maemaeMap.keySet()){
                System.out.println("[key]:" + i + " [value]:" + maemaeMap.get(i));
            }

            JMMultipleratioWrite jmMultipleratioWrite = new JMMultipleratioWrite();
            jmMultipleratioWrite.writeExcel(resultMap,maemaeMap, filePath, outputFileNm);
            System.out.println();

        } catch (Exception e){
            e.printStackTrace();
        }

    }


}



