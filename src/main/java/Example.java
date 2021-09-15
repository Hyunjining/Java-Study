import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

class Friend {
    String name, tel, state, nationality, job;

    Friend(String str1, String str2, String str3, String str4, String str5){
        this.name = str1;
        this.tel = str2;
        this.state = str3;
        this.nationality = str4;
        this.job = str5;
    }

}
public class Example {
    public static void main(String[] args) {

        String [] titles = {"name", "tel", "state", "nationality", "job"};
        String [] data = {
                "Kim,010-9169-0120,Seoul,Korea,Student",
                "Lee,010-5686-4435,Incheon,Korea,Office worker",
                "Park,010-4563-2343,Daejeon,Korea,Chef",
                "Lisa,010-4567-8765,Seoul,Korea,Barista",
                "Nana,010-3424-6453,Incheon,Korea,Writer",
                "Hi,010-5563-8767,Seoul,Korea,PD",
                "John,010-2343-9876,Gwangju,Korea,Saleman",
                "Sin,010-4563-9576,Daejeon,Korea,Developerr",
                "Hong,010-3645-2636,Dae-gu,Korea,Teacher",
                "Jin,010-2586-3452,Seong-Nam,Korea,Student",
                "Hyun,010-2342-3453,Busan,Korea,Developer",
                "Mi,010-3544-2264,Seong-Nam,Korea,Office worker",
                "Kyoung,010-2234-3645,Daejeon,Korea,Office worker",
                "An,010-2343-4463,Seoul,Korea,Developerr",
                "Young,010-2636-6789,Gwangju,Korea,Developer",
                "Jea,010-8758-9975,Seoul,Korea,Office worker",
                "Sang,010-5576-4656,Jeonju,Korea,Writer",
                "La,010-6756-4886,Busan,Korea,Developer",
                "Jang,010-5576-4435,Jeonju,Korea,Teacher",
                "Cho,010-5886-4485,Dae-gu,Korea,Student"
        };

        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("Friend Sheet");

        Row rowTitle = sheet.createRow(0);
        for (int i=0; i < titles.length; i++) {
            Cell cellTitle = rowTitle.createCell(i);
            cellTitle.setCellValue(titles[i]);
        }

        for (int i=0; i < data.length; i++) {
            String [] item = data[i].split(",");
            Row row = sheet.createRow(i+1);
            Cell cell = row.createCell(0);
            for(int j=0;j<5;j++){
                row.createCell(j).setCellValue(item[j]);
            }
        }



        int rows = sheet.getPhysicalNumberOfRows();
        for(int i=1;i<rows+1;i++){
            Row row = sheet.getRow(i);
            if(row != null){
                List<String> cellList = new ArrayList<String>();
                int cells = row.getPhysicalNumberOfCells();
                for(int c = 0; c<cells+1;c++){
                    Cell cell = row.getCell(c);
                    if(cell != null){
                        if(c==0){
                            cellList.add(String.valueOf(i-1));
                        }
                        else {
                            cellList.add(cell.getStringCellValue());
                        }
                    }
                    else{
                        cellList.add(" ");
                    }
                }
                for(String Data: cellList){
                    System.out.println(Data);
                }
            }

        }

        try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
            wb.write(fileOut);
        }
        catch(Exception e) {
        }

    }
}
