import java.io.*;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class XLSXREADER {
    public static void main(String[] args) {
        try {

            File excel = new File("/Users/deepanshu/Downloads/f.xlsx");
            FileInputStream fis = new FileInputStream(excel);
            XSSFWorkbook book = new XSSFWorkbook(fis);
            XSSFSheet sheet = book.getSheetAt(0);
            Integer startingRowNumber = 617;
            Integer endingRowNumber = 617;

            try {
                File myObj = new File("translated.txt");
                if (myObj.createNewFile()) {
                    System.out.println("File created: " + myObj.getName());
                } else {
                    System.out.println("File already exists.");
                }
            } catch (IOException e) {
                System.out.println("An error occurred.");
                e.printStackTrace();
            }

            FileWriter myWriter = new FileWriter("translated.txt");

          //  languageToTranslation.put("English", "Your PAN verification is completed. Please continue to verify Bank account.");
         //   englishToVernacular.put(languageToTranslation.get("English").toLowerCase(), languageToTranslation);
            for (int i=startingRowNumber; i<=endingRowNumber; i++) {
                myWriter.write("languageToTranslation = new HashMap<>();\n");
                XSSFRow row = sheet.getRow(i);
                Cell cell = row.getCell(0);
                Integer startColumnNumber =0;
                Integer endColumnNumber = 19;
                while(startColumnNumber< endColumnNumber) {
                    if(startColumnNumber==1) {
                        startColumnNumber+=3;
                        continue;
                    }
                    String header = sheet.getRow(0).getCell(startColumnNumber).getStringCellValue();
                    if(header.equals("Strings"))
                        header= "English";
                    String value = sheet.getRow(i).getCell(startColumnNumber).getStringCellValue().replaceAll("^\\s+|\\s+$", "");
                    myWriter.write("languageToTranslation.put"+"(\""+header+"\","+ " \""+value+"\");\n");
                    startColumnNumber++;
                }

                myWriter.write("englishToVernacular.put(languageToTranslation.get(\"English\").toLowerCase(), languageToTranslation);\n");
                myWriter.write("\n");

            }
            myWriter.close();



        }
        catch (Exception e){

        }

    }




}
