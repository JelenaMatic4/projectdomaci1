package D_22_Jelena_Matic_NS125QA6;

import com.github.javafaker.Faker;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ReadFile {
    public static void main(String[] args) {
        readExcel("domaci22.xlsx");
        try{
            writeExel("novatabela.xlsx");}
        catch (IOException e){
            System.out.println("Invalid file!");
        }
readExcel("novatabela.xlsx");
    }
public static void readExcel(String path){
        try{
            FileInputStream inputStream = new FileInputStream(new File(path));
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            for(int j=0; j<2; j++){
                XSSFRow row = sheet.getRow(j);
                for (int i = 0; i <2 ; i++) {
                    XSSFCell celle = row.getCell(i);
                    String imePrezime = celle.getStringCellValue();
                    System.out.print(imePrezime + " ");

                } System.out.println();

            }

        } catch (FileNotFoundException ex){
            System.out.println("File not find!");
        } catch (IOException e){
            e.printStackTrace();
        }catch(NullPointerException e){
        }

    }
    public static void writeExel (String fileName) throws IOException{
        Faker faker = new Faker();
        String name = faker.name().fullName(); // Miss Samanta Schmidt
        String firstName = faker.name().firstName(); // Emory
        String lastName = faker.name().lastName(); // Barton
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        for(int i =2; i<10; i++){
            XSSFRow row = sheet.createRow(i);
            for(int j=0; j<1; j++){
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(name);
                FileOutputStream fileOutputStream = new FileOutputStream(new File(fileName));
                workbook.write(fileOutputStream);
                fileOutputStream.close();
            }

        }

        try {
            FileInputStream inputStream = new FileInputStream(new File("novatabela.xlsx"));
            XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);
            XSSFSheet sheet1 = workbook.getSheet("Sheet1");
            for (int j = 0; j < 8; j++) {
                XSSFRow row = sheet1.getRow(j);
                for (int i = 0; i < 2; i++) {
                    XSSFCell celle = row.getCell(i);
                    String imeFake = celle.getStringCellValue();
                    System.out.print(imeFake + " ");

                }
                System.out.println();

            }
        }catch (FileNotFoundException ex){
                System.out.println("File not find!");
            } catch (IOException e){
                e.printStackTrace();
            }catch(NullPointerException e){
            }
         }


        }




