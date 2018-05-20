import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.concurrent.LinkedBlockingDeque;
import java.util.concurrent.LinkedBlockingQueue;

public class PeriodData {

    CompanyData[] companies;
    int maxRow;
    LocalDate firstDate;
    int period;
    private String path = "C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Matlab\\Data US\\US_Data_P";

    PeriodData(int period, int totalCompanies) throws IOException, InvalidFormatException {
        this.period = period;
        path = path + period;
        maxRow = getMaxRow(totalCompanies);
        firstDate = getFirstDate(totalCompanies);
    }

    private int getMaxRow(int totalCompanies) throws InvalidFormatException, IOException {
        int wbNum = 1;
        int sheetNum = 0;
        Integer[] rows = new Integer[totalCompanies];

        //Open the first Workbook
        //=================================================================================================================================
        XSSFWorkbook wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));

        for(int i=0;i<totalCompanies;i++){
            if(wbNum!=(int) Math.ceil((i+1)/100d) && i >= 99){ //True if current workbook unequals i / 100
                wbNum = (int) Math.ceil((i+1)/100d);//Sets workbook number new
                sheetNum = 0;
                wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));//define new workbook
            }

            rows[i] = wb.getSheetAt(sheetNum).getLastRowNum();//get the last row of each sheet
            sheetNum++;
        }

        return Collections.max(Arrays.asList(rows));//Find the highest number of rows
    }

    private LocalDate getFirstDate(int totalCompanies) throws InvalidFormatException, IOException {
        int wbNum = 1;
        int sheetNum = 0;
        LocalDate[] dates = new LocalDate[totalCompanies];

        XSSFWorkbook wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));

        for(int i=0;i<totalCompanies;i++){
            if(wbNum!=(int) Math.ceil((i+1)/100d) && i >= 99){
                wbNum = (int) Math.ceil((i+1)/100d);
                sheetNum = 0;
                wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));
            }
            System.out.println("Workbook: " + wbNum +" |Sheet: " + sheetNum);
            dates[i] = wb.getSheetAt(sheetNum).getRow(1).getCell(0).getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();// get first date of each sheet
            sheetNum++;
        }
        Arrays.sort(dates);//sort dates

        return dates[0]; //return earliest date
    }

    public LinkedBlockingQueue getQueuedSheets(int totalCompanies) throws IOException, InvalidFormatException, InterruptedException {
        int wbNum = 1;
        int sheetNum = 0;
        LinkedBlockingQueue<Sheet> sheetQueue = new LinkedBlockingQueue<>(3000); //Queue for storing sheets for later retrieval

        XSSFWorkbook wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));

        for(int i=0;i<totalCompanies;i++){
            if(wbNum!=(int) Math.ceil((i+1)/100d) && i >= 99){
                wbNum = (int) Math.ceil((i+1)/100d);
                sheetNum = 0;
                wb = new XSSFWorkbook(OPCPackage.open(new File(path + "_" + wbNum + ".xlsx")));
            }

            if(checkPeriod(wb.getSheetAt(sheetNum))){ //if Sheet has the necessary number of rows and starts with the first date then true
                sheetQueue.put(wb.getSheetAt(sheetNum));
            }
            sheetNum++;
        }
        return sheetQueue;
    }

    public void getCompanyData() throws IOException, InvalidFormatException {
        // XSSFWorkbook, File
        OPCPackage pkg = OPCPackage.open(new File("C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Matlab\\Data US\\US_Data_P1_1.xlsx"));
        XSSFWorkbook wb = new XSSFWorkbook(pkg);
        Sheet sheet = wb.getSheetAt(0);
        CompanyData comp = new CompanyData(sheet);
        GlobalData globalData = new GlobalData(sheet);


        //Daten Import
        //========================================================================================================================================

        for(int i=1;i<sheet.getLastRowNum()+1;i++) {
            comp.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
            globalData.dateAdd(sheet.getRow(i).getCell(0).getDateCellValue());
        }

        //Add Company specific Returns
        comp.addReturns("Mid");
        comp.addReturns( "Bid");
        comp.addReturns( "Ask");

        //Add Company specific Market Data
        comp.addMarketData("MktCap");
        comp.addMarketData( "Volume");

        //Global Data
        globalData.addMarketData("S&P");
        globalData.addMarketData( "Russel");
        globalData.addMarketData( "RiskFree");

        //Company Data
        comp.addCompanyData( "PB");
        comp.addCompanyData( "PE");
        comp.addCompanyData( "PS");


        //Lagged Returns
        comp.addLaggedReturns( "oneWeek");
        comp.addLaggedReturns( "oneMonth");
        comp.addLaggedReturns( "threeMonth");
        comp.addLaggedReturns( "sixMonth");

        //Volatility
        comp.addVolatility( "oneWeek");
        comp.addVolatility( "oneMonth");
        comp.addVolatility( "threeMonth");
        comp.addVolatility( "sixMonth");

        comp.writeData("C:\\Users\\kimwa\\OneDrive\\Documents\\Codes\\Java\\US-OutPut\\test.csv", true);
    }

    private Boolean checkPeriod(Sheet sheet){
        return (firstDate.equals(sheet.getRow(1).getCell(0).getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate()) &&
                sheet.getLastRowNum()>=((2/3)*maxRow));
    }
}
