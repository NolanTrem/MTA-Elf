/*MTA Data Search reads txt files of weekly MTA Data, searches for specific stations, and exports this data to Excel.
* Pulls key data in the MTA designated format.
* MTA turnstile data can be found at: http://web.mta.info/developers/turnstile.html
* Nolan Tremelling, Columbia University 2020
*/
import java.io.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.xml.crypto.Data;
import java.util.*;


public class Main {
    public static void main(String[] args)throws Exception {
        //Following lines can be customized based on user needs.
        String[] columns = {"Station", "Linename", "Date", "Time", "Entries", "Exits", "Change in Time", "Change in Entries", "Change in Exits"};
        File file = new File("C:\\Users\\nolan\\Desktop\\COVID-19\\MTA Data Search\\turnstile_200620.txt");
        //Necessary declarations to read file
        BufferedReader br = new BufferedReader(new FileReader(file));
        String st;
        String station = "72 ST";    //Station name per MTA stylization
        List<DataLine> DataList = new ArrayList<>();

        while ((st = br.readLine()) != null){
            if (st.contains(station)) {     // Change this string to search for other stations.
                String tokens[] = st.split(","); //MTA compiles data with separations by comma.
                DataList.add(new DataLine(tokens[3], tokens[4], tokens[6], tokens[7], tokens[9], tokens[10])); //Target data

                //Sets up Excel workbook.
                Workbook workbook = new XSSFWorkbook();
                CreationHelper createHelper = workbook.getCreationHelper();
                Sheet sheet = workbook.createSheet(station);
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setFontHeightInPoints((short) 14);
                headerFont.setColor(IndexedColors.ROYAL_BLUE.getIndex());
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
                    row.createCell(6).setCellValue("hello");
                    int entriesLow = 2;
                    int entriesHigh = 3;
                    int exitsLow = 2;
                    int exitsHigh = 3;
                    int asdf = 0;
                    while (asdf < rowNum){
                        entriesLow++;
                        entriesHigh++;
                        exitsLow++;
                        exitsHigh++;
                        sheet.autoSizeColumn(asdf);
                        row.createCell(7).setCellFormula("E" + entriesHigh + " - E" + entriesLow);
                    }
                    row.createCell(8).setCellValue("hello");
                }


                //Change the following line to change the title of the Excel file.
                FileOutputStream fileOut = new FileOutputStream(station + ".xlsx");
                workbook.write(fileOut);
                fileOut.close();
                workbook.close();
            }
        }
    }
}
