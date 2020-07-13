import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.crypto.Data;
import java.util.*;

class DataLine {
    public String station;
    public String getStation() {
        return this.station;
    }
    public String linename;
    public String getLinename() {
        return this.linename;
    }
    public String date;
    public String getDate(){
        return this.date;
    }
    public String time;
    public String getTime(){
        return this.time;
    }
    public String entries;
    public String getEntries(){
        return this.entries;
    }
    public String exits;
    public String getExits(){
        return this.exits;
    }

    public DataLine(String Station, String Linename, String Date, String Time, String Entries, String Exits) {
        this.station = Station;
        this.linename = Linename;
        this.date = Date;
        this.time = Time;
        this.entries = Entries;
        this.exits = Exits;
        //System.out.println(station + "\t" + linename + "\t" + date + "\t" + time + "\t" + entries + "\t" + exits);

    }
}

public class Main {
    public static void main(String[] args)throws Exception {
        String[] columns = {"Station", "Linename", "Date", "Time", "Entries", "Exits"};
        File file = new File("C:\\Users\\nolan\\Desktop\\COVID-19\\MTA Data Search\\turnstile_200620.txt");
        BufferedReader br = new BufferedReader(new FileReader(file));
        String st;
        List<DataLine> DataList = new ArrayList<>();

        while ((st = br.readLine()) != null){
            if (st.contains("72 ST")) {     // Change this string to search for other stations
                String tokens[] = st.split(",");
                DataList.add(new DataLine(tokens[3], tokens[4], tokens[6], tokens[7], tokens[9], tokens[10]));

                Workbook workbook = new XSSFWorkbook();
                CreationHelper createHelper = workbook.getCreationHelper();
                Sheet sheet = workbook.createSheet("TITLE");
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setFontHeightInPoints((short) 14);
                headerFont.setColor(IndexedColors.RED.getIndex());
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFont(headerFont);
                Row headerRow = sheet.createRow(0);

                for(int i = 0; i < columns.length; i++){
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(columns[i]);
                    cell.setCellStyle(headerCellStyle);
                }

                int rowNum = 1;
                for (DataLine dataLine: DataList){
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(dataLine.getStation());
                    row.createCell(1).setCellValue(dataLine.getLinename());
                    row.createCell(2).setCellValue(dataLine.getDate());
                    row.createCell(3).setCellValue(dataLine.getTime());
                    row.createCell(4).setCellValue(dataLine.getEntries());
                    row.createCell(5).setCellValue(dataLine.getExits());
                }

                for(int i = 0; i < columns.length; i++){
                    sheet.autoSizeColumn(i);
                }

                FileOutputStream fileOut = new FileOutputStream("poi-generated.xlsx");
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
            }
        }
    }


}
