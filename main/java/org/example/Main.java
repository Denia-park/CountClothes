package org.example;

import com.opencsv.CSVWriter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.text.NumberFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class Main {
    static final String PROGRAM_VERSION = "Version : 1.2 , UpdateDate : 22년 9월 29일";
    static String fileNameCSV = getStringOfNowLocalDateTime();
    static final int PRODUCT_NAME_CELL_INDEX = 1; //B [0부터 시작임.]
    static final int PRODUCT_OPTION_CELL_INDEX = 2; //C [0부터 시작임.]
    static final int PRODUCT_QUANTITY_CELL_INDEX = 3; //D [0부터 시작임.]

    public static void main(String[] args) {
        System.out.println("옷의 수량을 세는 프로그램을 시작합니다. [ " + PROGRAM_VERSION + " ]");

        Map<String, Integer> map = new TreeMap<>();

        Scanner sc = new Scanner(System.in); // 사용자로부터 데이터를 받기 위한 Scanner

        String path = System.getProperty("user.dir") + "\\"; //현재 작업 경로
        String fileName = "countClothes.xlsx"; //파일명 설정

        XSSFSheet sheetDataFromExcel = readExcel(path, fileName); //엑셀 파일 Read
        if (sheetDataFromExcel == null) { //파일을 못 읽어오면 종료.
            System.out.println("파일을 찾지 못했으므로 프로그램을 종료 합니다.");

            System.out.println("Enter 를 치면 정상 종료됩니다.");
            sc.nextLine(); //프로그램 종료 전 Holding
            return; //프로그램 종료
        }

        //행 갯수 가져오기
        int rows = sheetDataFromExcel.getPhysicalNumberOfRows();

        XSSFRow row = sheetDataFromExcel.getRow(0); //Title Row 가져오기
        int cells = row.getPhysicalNumberOfCells(); //Title Cell 수 가져오기
        String[][] dataBufferArr = new String[2][cells]; //행을 읽어서 저장해둘 배열을 생성
        int currentSaveOrder = 0; //dataBufferArr 에서 몇번째 배열인지 알려줄 인자
        NumberFormat f = NumberFormat.getInstance(); //엑셀에서 NumberFormat이 나왔을때 저장할 수 있게 생성함
        f.setGroupingUsed(false);	//지수로 안나오게 설정

        //반드시 "행(row)"을 읽고 "열(cell)"을 읽어야함 ..
        //rowIndex = 0 => Title
        for(int rowIndex = 1 ; rowIndex < rows ; rowIndex++) {
            row = sheetDataFromExcel.getRow(rowIndex);

            for (int i = 0; i < cells; i++) {
                XSSFCell cell = row.getCell(i);
                dataBufferArr[currentSaveOrder][i] = readCell(cell,f);
            }

            //상품명 + 옵션 을 합친 String 을 Key로 사용
            String mapKey = dataBufferArr[currentSaveOrder][PRODUCT_NAME_CELL_INDEX]
                    + " / "
                    + dataBufferArr[currentSaveOrder][PRODUCT_OPTION_CELL_INDEX];

            //수량이 String 으로 되어있으므로 Integer 파싱 후 사용
            int clothesQuantity = Integer.parseInt(dataBufferArr[currentSaveOrder][PRODUCT_QUANTITY_CELL_INDEX]);

            //map 에 내용들을 저장
            map.put(mapKey, map.getOrDefault(mapKey, 0) + clothesQuantity);

            currentSaveOrder ^= 1; //dataBufferArr 에 저장할 순서 변경 0 -> 1 , 1 -> 0 : XOR을 사용했다.
        }

        // map에 저장한 내용들을 CSV 파일에 저장하기
        writeDataToCSV(path, map);

        System.out.println("작업이 완료되었습니다.");

        System.out.println("Enter 를 치면 정상 종료됩니다.");
        sc.nextLine(); //프로그램 종료 전 Holding
    }

    private static void writeDataToCSV(String path, Map<String, Integer> map) {
        File file = new File(path, fileNameCSV);
        int clothesTotalQuantity = 0;

        try (
                FileOutputStream fos = new FileOutputStream(file,true);
                OutputStreamWriter osw = new OutputStreamWriter(fos, StandardCharsets.UTF_8);
                CSVWriter writer = new CSVWriter(osw)
        ) {
            //제목 저장
            String[] title = {
                    "상품명",
                    "옵션",
                    "수량",
            };
            writer.writeNext(title,false);

            //entrySet 을 돌면서 map의 내용을 읽어 들인 후 csv 에 출력
            for (Map.Entry<String, Integer> entrySet : map.entrySet()) {
                int mapValue = entrySet.getValue();
                clothesTotalQuantity += mapValue;

                String[] productInfo = entrySet.getKey().split("/");
                String productName = productInfo[0];
                String productOption = productInfo[1];

                String[] data = {productName, productOption, String.valueOf(mapValue)};

                writer.writeNext(data, false);
            }

            //전체 수량도 출력
            String[] totalSum = {
                    "※전체 수량※ : ",
                    String.valueOf(clothesTotalQuantity),
            };
            writer.writeNext(totalSum,false);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static XSSFSheet readExcel(String path, String fileName){
        try {
            FileInputStream file = new FileInputStream(path + fileName);
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            return workbook.getSheetAt(0); // 첫번째 시트만 사용
        } catch(IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static String readCell(XSSFCell cell, NumberFormat f) {
        String tempValue = "";
        if(cell != null){
            //타입 체크
            switch(cell.getCellType()) {
                case STRING:
                    tempValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    tempValue = f.format(cell.getNumericCellValue())+"";
                    break;
                case BLANK:
                    tempValue = "";
                    break;
                case ERROR:
                    tempValue = cell.getErrorCellValue()+"";
                    break;
            }
            return tempValue;
        }
        else
            throw new RuntimeException("Cell Read 중 NPE 발생함");
    }
    private static String getStringOfNowLocalDateTime() {
        // 현재 날짜/시간
        LocalDateTime now = LocalDateTime.now(); // 2021-06-17T06:43:21.419878100

        // 포맷팅
        String formatedNow = now.format(DateTimeFormatter.ofPattern("yyMMdd_HH_mm_ss")); // 220628_02_38_02

        return "Count Clothes TXT_" + formatedNow + ".txt"; //Ex) CSV_220628_02_38_02

    }
}
