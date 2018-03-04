import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Arrays;
import java.util.Collections;

public class PeriodData {

    CompanyData[] companies;
    int maxRow;
    LocalDate firstDate;
    String Path;

    PeriodData(String Path, String Period) throws InvalidFormatException, IOException {
        this.Path = Path;

        OPCPackage pkg = OPCPackage.open(new File("C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Matlab\\Data US\\US_Data_P" + Period + "_1.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(pkg);

        int numSheets = wb.getNumberOfSheets();

        Integer[] rows = new Integer[numSheets];
        LocalDate[] startDate = new LocalDate[numSheets];
        //find MaxNumber of dates

        for(int i=0;i<numSheets;i++){
            Sheet sheet = wb.getSheetAt(i);
            rows[i] = sheet.getLastRowNum();
            startDate[i] = sheet.getRow(1).getCell(0).getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            sheet = null;
        }

        Arrays.sort(rows, Collections.reverseOrder());
        maxRow = rows[0];
        rows = null;

        Arrays.sort(startDate);
        firstDate = startDate[0];
        startDate = null;

        pkg = null;
        wb = null;
    }

    Boolean checkPeriod(Sheet sheet){
        return (firstDate.equals(sheet.getRow(1).getCell(0).getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate()) &&
                sheet.getLastRowNum()>=((2/3)*maxRow));
    }
}
